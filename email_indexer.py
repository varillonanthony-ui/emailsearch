"""
email_indexer.py – Synchronisation emails Office 365 via Microsoft Graph API.

Stratégie v5 :
  • $top=500 (max autorisé) → 20 appels pour 10 000 emails au lieu de 100+
  • Corps (body) NON téléchargé pendant la sync bulk → réponses 50x plus légères
    Le corps est chargé à la demande quand l'utilisateur clique "Voir le contenu complet"
  • bodyPreview (255 car.) indexé → suffisant pour la recherche par mot-clé
  • Checkpoint intra-dossier : le nextLink en cours est sauvegardé toutes les 500 emails
    Si Streamlit est interrompu, la sync reprend exactement là où elle s'est arrêtée
  • Gestion 410 Gone, throttling 429, 5xx, timeout 60s
"""

import re
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable
import requests

from database import Database

GRAPH = "https://graph.microsoft.com/v1.0"

# Corps exclu du select bulk → réponses légères, sync rapide
# Le corps est chargé à la demande via get_email_body()
MSG_SELECT = (
    "id,subject,from,sender,toRecipients,ccRecipients,"
    "bodyPreview,receivedDateTime,sentDateTime,"
    "hasAttachments,isRead,importance,conversationId,webUrl"
)

# $top=500 = maximum autorisé par Graph pour /messages/delta
PAGE_SIZE = 500

WELL_KNOWN = [
    "inbox", "sentitems", "deleteditems", "drafts",
    "junkemail", "archive", "clutter", "outbox",
]


@dataclass
class SyncResult:
    total_folders:  int = 0
    folders_done:   int = 0
    folders_skip:   int = 0
    emails_new:     int = 0
    emails_updated: int = 0
    emails_deleted: int = 0
    errors:         list[str] = field(default_factory=list)

    @property
    def emails_total(self) -> int:
        return self.emails_new + self.emails_updated


def _clean(text: str | None) -> str:
    if not text:
        return ""
    return re.sub(r"\s+", " ", text.replace("\x00", "")).strip()


class _GoneError(Exception):
    """HTTP 410 – delta token expiré."""


class EmailIndexer:

    def __init__(self, access_token: str, user_id: str):
        self.db = Database(user_id)
        self._headers = {
            "Authorization": f"Bearer {access_token}",
        }

    # ── HTTP ──────────────────────────────────────────────────────────────────

    def _get(self, url: str, params: dict | None = None) -> dict:
        for attempt in range(6):
            try:
                r = requests.get(
                    url, headers=self._headers,
                    params=params, timeout=60,
                )
            except requests.RequestException as exc:
                if attempt == 5:
                    raise RuntimeError(f"Réseau : {exc}") from exc
                time.sleep(2 ** attempt)
                continue

            if r.status_code == 410:
                raise _GoneError("Delta token expiré (410)")
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", 20))
                time.sleep(wait)
                continue
            if r.status_code == 401:
                raise PermissionError("Token Microsoft expiré – reconnectez-vous.")
            if r.status_code in (500, 502, 503, 504):
                time.sleep(2 ** attempt)
                continue
            if not r.ok:
                raise RuntimeError(f"HTTP {r.status_code} : {r.text[:300]}")
            return r.json()

        raise RuntimeError(f"Abandon après 6 tentatives : {url[:80]}")

    # ── Corps à la demande ────────────────────────────────────────────────────

    def get_email_body(self, email_id: str) -> str:
        """Télécharge le corps d'un email unique (appelé à la demande depuis l'UI)."""
        try:
            headers = {
                **self._headers,
                "Prefer": 'outlook.body-content-type="text"',
            }
            r = requests.get(
                f"{GRAPH}/me/messages/{email_id}",
                headers=headers,
                params={"$select": "body"},
                timeout=30,
            )
            if r.ok:
                return _clean(r.json().get("body", {}).get("content", ""))
        except Exception:
            pass
        return ""

    # ── Dossiers ──────────────────────────────────────────────────────────────

    def _fetch_folders(self) -> list[dict]:
        seen:    set[str]   = set()
        folders: list[dict] = []

        def _save(f: dict, path: str, parent_id: str | None):
            fid = f["id"]
            if fid in seen:
                return
            seen.add(fid)
            rec = {
                "id":                fid,
                "name":              f["displayName"],
                "parent_folder_id":  parent_id or "",
                "total_item_count":  f.get("totalItemCount", 0),
                "unread_item_count": f.get("unreadItemCount", 0),
                "display_path":      path,
                "last_sync":         None,
            }
            folders.append(rec)
            self.db.upsert_folder(rec)

        # 1. Dossiers système garantis
        for wk in WELL_KNOWN:
            try:
                f = self._get(f"{GRAPH}/me/mailFolders/{wk}")
                _save(f, f["displayName"], None)
            except Exception:
                pass

        # 2. Parcours récursif complet
        def _recurse(parent_id: str | None, parent_path: str):
            url = (
                f"{GRAPH}/me/mailFolders/{parent_id}/childFolders"
                if parent_id else f"{GRAPH}/me/mailFolders"
            )
            params = {"$top": "100"}
            while url:
                data   = self._get(url, params if "?" not in url else None)
                for f in data.get("value", []):
                    path = f"{parent_path}/{f['displayName']}" if parent_path else f["displayName"]
                    _save(f, path, parent_id)
                    if f.get("childFolderCount", 0) > 0:
                        _recurse(f["id"], path)
                url    = data.get("@odata.nextLink")
                params = None

        _recurse(None, "")
        return folders

    # ── Parsing ───────────────────────────────────────────────────────────────

    @staticmethod
    def _parse(msg: dict, folder_id: str, folder_name: str) -> dict:
        ea = (msg.get("from") or msg.get("sender") or {}).get("emailAddress", {})

        def _addrs(key: str) -> str:
            return "; ".join(
                r["emailAddress"].get("address", "")
                for r in msg.get(key, [])
                if r.get("emailAddress", {}).get("address")
            )

        return {
            "id":                msg["id"],
            "folder_id":         folder_id,
            "folder_name":       folder_name,
            "subject":           _clean(msg.get("subject")) or "(Sans objet)",
            "sender_name":       _clean(ea.get("name", "")),
            "sender_email":      ea.get("address", ""),
            "recipients":        _clean(_addrs("toRecipients")),
            "cc":                _clean(_addrs("ccRecipients")),
            "body_preview":      _clean(msg.get("bodyPreview", "")),
            "body":              "",          # chargé à la demande
            "received_datetime": msg.get("receivedDateTime", ""),
            "sent_datetime":     msg.get("sentDateTime", ""),
            "has_attachments":   1 if msg.get("hasAttachments") else 0,
            "is_read":           1 if msg.get("isRead") else 0,
            "importance":        msg.get("importance", "normal"),
            "conversation_id":   msg.get("conversationId", ""),
            "web_link":          msg.get("webUrl", "") or msg.get("webLink", ""),
            "indexed_at":        datetime.utcnow().isoformat(),
        }

    # ── Sync d'un dossier ─────────────────────────────────────────────────────

    def _sync_folder(
        self,
        folder_id:   str,
        folder_name: str,
        force_full:  bool = False,
        on_progress: Callable | None = None,
    ) -> tuple[bool, int, int, int]:
        """
        Retourne (was_incremental, new, updated, deleted).

        Checkpoint intra-dossier :
        - Le nextLink en cours est sauvegardé toutes les PAGE_SIZE emails
        - Si interrompu, reprend depuis ce nextLink
        - Le deltaLink final remplace le nextLink checkpoint en fin de sync
        """
        delta_key     = f"delta_{folder_id}"
        cursor_key    = f"cursor_{folder_id}"  # nextLink checkpoint intra-dossier

        # Détermine l'URL de départ
        if force_full:
            # Reset complet de ce dossier
            self.db.set_sync_state(delta_key, "")
            self.db.set_sync_state(cursor_key, "")
            was_inc = False
            start_url = None
        else:
            delta_link  = self.db.get_sync_state(delta_key) or ""
            cursor_link = self.db.get_sync_state(cursor_key) or ""

            if delta_link:
                # Mode incrémental : reprend depuis le delta token
                was_inc   = True
                start_url = delta_link
            elif cursor_link:
                # Reprise après interruption dans un dossier
                was_inc   = False
                start_url = cursor_link
            else:
                was_inc   = False
                start_url = None

        # URL initiale (si pas de checkpoint/delta)
        if start_url is None:
            start_url = f"{GRAPH}/me/mailFolders/{folder_id}/messages/delta"

        # Paramètres uniquement pour l'appel initial (pas inclus dans nextLink/deltaLink)
        initial_params = {"$select": MSG_SELECT, "$top": str(PAGE_SIZE)}

        # Chargement bulk des IDs existants (1 seule requête SQL)
        existing: set[str] = self.db.get_email_ids_for_folder(folder_id)

        new_n = upd_n = del_n = 0
        batch: list[dict] = []
        final_delta: str | None = None
        url = start_url

        # Détermine si on passe les params initiaux (URL propre sans query string)
        use_initial_params = "?" not in url

        try:
            while url:
                data = self._get(
                    url,
                    initial_params if use_initial_params else None,
                )
                use_initial_params = False  # seulement pour le premier appel

                for msg in data.get("value", []):
                    if msg.get("@removed"):
                        self.db.delete_email(msg["id"])
                        existing.discard(msg["id"])
                        del_n += 1
                        continue

                    parsed = self._parse(msg, folder_id, folder_name)
                    if parsed["id"] in existing:
                        upd_n += 1
                    else:
                        new_n += 1
                        existing.add(parsed["id"])
                    batch.append(parsed)

                    if len(batch) >= PAGE_SIZE:
                        self.db.upsert_emails_batch(batch)
                        batch = []
                        # Checkpoint intra-dossier (nextLink actuel)
                        if url:
                            self.db.set_sync_state(cursor_key, url)
                        if on_progress:
                            on_progress(new_n, upd_n, del_n)

                next_link  = data.get("@odata.nextLink")
                delta_link = data.get("@odata.deltaLink")

                if delta_link:
                    final_delta = delta_link
                    url = None
                elif next_link:
                    url = next_link
                else:
                    # Ni nextLink ni deltaLink : fin inattendue, on arrête proprement
                    url = None

        except _GoneError:
            # Delta expiré → full sync automatique pour ce dossier
            if batch:
                self.db.upsert_emails_batch(batch)
            self.db.set_sync_state(delta_key, "")
            self.db.set_sync_state(cursor_key, "")
            return self._sync_folder(folder_id, folder_name,
                                     force_full=True, on_progress=on_progress)

        # Flush du dernier batch
        if batch:
            self.db.upsert_emails_batch(batch)

        # Sauvegarde finale
        if final_delta:
            self.db.set_sync_state(delta_key, final_delta)
            self.db.set_sync_state(cursor_key, "")  # efface le checkpoint intermédiaire

        self.db.set_sync_state(f"folder_done_{folder_id}", datetime.utcnow().isoformat())
        return was_inc, new_n, upd_n, del_n

    # ── Orchestration ─────────────────────────────────────────────────────────

    def sync(
        self,
        force_full: bool = False,
        on_status:  Callable[[str, SyncResult], None] | None = None,
    ) -> SyncResult:
        result = SyncResult()

        if on_status:
            on_status("📁 Récupération des dossiers…", result)

        folders = self._fetch_folders()
        result.total_folders = len(folders)

        if on_status:
            on_status(f"📁 {len(folders)} dossier(s) trouvé(s)", result)

        for i, folder in enumerate(folders):
            fid   = folder["id"]
            fname = folder["display_path"]
            count = folder.get("total_item_count", 0)
            label = f"[{i+1}/{len(folders)}] {fname} ({count} emails)"

            # Checkpoint : dossier déjà entièrement synchronisé ?
            if (
                not force_full
                and self.db.get_sync_state(f"delta_{fid}")
                and not self.db.get_sync_state(f"cursor_{fid}")
                # cursor vide = pas de sync interrompue en cours
            ):
                result.folders_skip += 1
                if on_status:
                    on_status(f"⏭ Déjà à jour : {fname}", result)
                continue

            if on_status:
                on_status(f"🔄 {label}", result)

            try:
                inc, n, u, d = self._sync_folder(
                    folder_id=fid,
                    folder_name=folder["name"],
                    force_full=force_full,
                    on_progress=lambda nw, up, dl, r=result, lbl=label: (
                        on_status(
                            f"🔄 {lbl} — {r.emails_new + nw + r.emails_updated + up:,} traités",
                            r
                        ) if on_status else None
                    ),
                )
            except PermissionError:
                raise
            except Exception as exc:
                result.errors.append(f"{fname} : {exc}")
                if on_status:
                    on_status(f"⚠️ Erreur {fname} : {exc}", result)
                continue

            result.emails_new     += n
            result.emails_updated += u
            result.emails_deleted += d
            result.folders_done   += 1

            if on_status:
                mode = "incrémental" if inc else "complet"
                on_status(
                    f"✅ {fname} [{mode}] → +{n} / ✏{u} / 🗑{d}",
                    result,
                )

        self.db.set_sync_state("last_full_sync", datetime.utcnow().isoformat())
        return result