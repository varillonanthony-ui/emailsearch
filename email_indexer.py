"""
email_indexer.py v6 — Approche simple et fiable

Sync complète  : GET /mailFolders/{id}/messages?$top=1000
                 Pagine via @odata.nextLink jusqu'à épuisement total.
                 Vérifie à la fin que le nombre indexé ≈ totalItemCount.

Sync incrémentale : même endpoint avec $filter sur receivedDateTime
                    + récupère les emails modifiés récemment.

Avantages vs delta :
  - Pas de deltaLink prématuré
  - Vérifiable : on sait combien d'emails le dossier doit avoir
  - Reprise sur interruption via checkpoint de page (skipToken)
  - Aucun état corrompu possible
"""

import re
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone, timedelta
from typing import Callable
import requests

from database import Database

GRAPH      = "https://graph.microsoft.com/v1.0"
# Avec le corps texte inclus, on réduit la taille de page pour éviter les timeouts.
# 50 messages × ~5 Ko de corps texte = ~250 Ko par requête, très raisonnable.
PAGE_SIZE  = 50

# body inclus dans la sync bulk (format texte brut, ~2-5 Ko/email)
# Indispensable pour trouver les mots-clés dans le corps complet.
# Le format HTML est exclu ($top=1000 resterait dans les limites)
MSG_SELECT = (
    "id,subject,from,sender,toRecipients,ccRecipients,"
    "bodyPreview,body,receivedDateTime,sentDateTime,"
    "hasAttachments,isRead,importance,conversationId,webLink"
)

WELL_KNOWN = [
    "inbox", "sentitems", "deleteditems", "drafts",
    "junkemail", "archive", "clutter", "outbox",
]


@dataclass
class FolderSyncInfo:
    name:         str
    expected:     int   # totalItemCount selon Graph
    indexed:      int   # emails en base après sync
    new:          int = 0
    updated:      int = 0


@dataclass
class SyncResult:
    total_folders:   int = 0
    folders_done:    int = 0
    folders_skip:    int = 0
    emails_new:      int = 0
    emails_updated:  int = 0
    warnings:        list[str] = field(default_factory=list)
    errors:          list[str] = field(default_factory=list)
    folder_details:  list[FolderSyncInfo] = field(default_factory=list)


def _clean(text: str | None) -> str:
    if not text:
        return ""
    return re.sub(r"\s+", " ", text.replace("\x00", "")).strip()


class EmailIndexer:

    def __init__(self, access_token: str, user_id: str):
        self.db = Database(user_id)
        self._headers = {
            "Authorization": f"Bearer {access_token}",
            # Demande le corps en texte brut (3-10x plus léger que HTML)
            "Prefer": 'outlook.body-content-type="text"',
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
                raise RuntimeError(f"HTTP {r.status_code} sur {url[:80]}: {r.text[:200]}")
            return r.json()

        raise RuntimeError(f"Abandon après 6 tentatives : {url[:80]}")

    # ── Corps à la demande ────────────────────────────────────────────────────

    def get_email_body(self, email_id: str) -> str:
        """
        Télécharge le corps HTML d'un email (appelé à la demande depuis l'UI).
        Retourne le HTML brut tel que renvoyé par Outlook.
        """
        try:
            r = requests.get(
                f"{GRAPH}/me/messages/{email_id}",
                headers={
                    **self._headers,
                    # Demande le corps en HTML (format natif Outlook)
                    "Prefer": 'outlook.body-content-type="html"',
                },
                params={"$select": "body"},
                timeout=30,
            )
            if r.ok:
                data    = r.json()
                content = data.get("body", {}).get("content", "")
                return content.replace("", "") if content else ""
        except Exception:
            pass
        return ""

    def list_attachments(self, email_id: str) -> list[dict]:
        """
        Liste les PJ d'un email (métadonnées uniquement, sans téléchargement).
        Exclut les images inline déjà intégrées dans le corps HTML.
        """
        try:
            r = requests.get(
                f"{GRAPH}/me/messages/{email_id}/attachments",
                headers=self._headers,
                params={"$select": "id,name,contentType,size,isInline"},
                timeout=30,
            )
            if r.ok:
                return [
                    {
                        "id":          a["id"],
                        "name":        a.get("name", "Sans nom"),
                        "contentType": a.get("contentType", "application/octet-stream"),
                        "size":        a.get("size", 0),
                        "isInline":    a.get("isInline", False),
                    }
                    for a in r.json().get("value", [])
                    if not a.get("isInline", False)
                ]
        except Exception:
            pass
        return []

    def get_attachment_content(self, email_id: str, attachment_id: str) -> tuple[str, str]:
        """
        Récupère le contenu base64 d'une PJ via Graph API.
        Les octets ne transitent JAMAIS par le disque utilisateur :
        Graph → base64 → mémoire Python → rendu navigateur.
        Retourne (base64_string, content_type).
        """
        try:
            r = requests.get(
                f"{GRAPH}/me/messages/{email_id}/attachments/{attachment_id}",
                headers=self._headers,
                params={"$select": "contentBytes,contentType,name"},
                timeout=60,
            )
            if r.ok:
                data = r.json()
                return data.get("contentBytes", ""), data.get("contentType", "")
        except Exception:
            pass
        return "", ""


    # ── Dossiers ──────────────────────────────────────────────────────────────

    def _fetch_folders(self) -> list[dict]:
        """Découverte exhaustive : well-known + récursif complet."""
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

        for wk in WELL_KNOWN:
            try:
                f = self._get(f"{GRAPH}/me/mailFolders/{wk}")
                _save(f, f["displayName"], None)
            except Exception:
                pass

        def _recurse(parent_id: str | None, parent_path: str):
            url    = (f"{GRAPH}/me/mailFolders/{parent_id}/childFolders"
                      if parent_id else f"{GRAPH}/me/mailFolders")
            params = {"$top": "100", "$includeHiddenFolders": "true"}
            while url:
                data   = self._get(url, params if "?" not in url else None)
                for f in data.get("value", []):
                    path = (f"{parent_path}/{f['displayName']}"
                            if parent_path else f["displayName"])
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

        def _addrs(key):
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
            "body":              _clean((msg.get("body") or {}).get("content", "")),
            "received_datetime": msg.get("receivedDateTime", ""),
            "sent_datetime":     msg.get("sentDateTime", ""),
            "has_attachments":   1 if msg.get("hasAttachments") else 0,
            "is_read":           1 if msg.get("isRead") else 0,
            "importance":        msg.get("importance", "normal"),
            "conversation_id":   msg.get("conversationId", ""),
            "web_link":          msg.get("webLink", ""),
            "indexed_at":        datetime.utcnow().isoformat(),
        }

    # ── Sync d'un dossier ─────────────────────────────────────────────────────

    def _sync_folder(
        self,
        folder_id:    str,
        folder_name:  str,
        expected:     int,
        since_dt:     str | None = None,   # None = full, sinon ISO date
        on_progress:  Callable | None = None,
    ) -> tuple[int, int]:
        """
        Pagine à travers TOUS les messages d'un dossier.
        Checkpoint : le nextLink courant est sauvegardé en base toutes les 1000 emails
        pour pouvoir reprendre en cas d'interruption.

        Retourne (new, updated).
        """
        cursor_key = f"cursor2_{folder_id}"
        saved_cursor = self.db.get_sync_state(cursor_key) or ""

        # Paramètres du premier appel
        base_url = f"{GRAPH}/me/mailFolders/{folder_id}/messages"
        if saved_cursor:
            # Reprise depuis un checkpoint intra-dossier
            url    = saved_cursor
            params = None
        else:
            url    = base_url
            params = {"$select": MSG_SELECT, "$top": str(PAGE_SIZE)}
            if since_dt:
                # Incrémental : seulement les emails plus récents que la dernière sync
                params["$filter"] = f"receivedDateTime ge {since_dt}"

        existing: set[str] = self.db.get_email_ids_for_folder(folder_id)
        new_n = upd_n = 0
        batch: list[dict] = []
        page_count = 0

        while url:
            data = self._get(url, params)
            params = None   # params seulement pour le premier appel

            for msg in data.get("value", []):
                parsed = self._parse(msg, folder_id, folder_name)
                if parsed["id"] in existing:
                    upd_n += 1
                else:
                    new_n += 1
                    existing.add(parsed["id"])
                batch.append(parsed)

            # Flush batch
            if len(batch) >= PAGE_SIZE:
                self.db.upsert_emails_batch(batch)
                batch = []

            page_count += 1
            url = data.get("@odata.nextLink")

            # Checkpoint : sauvegarde le nextLink toutes les pages
            if url:
                self.db.set_sync_state(cursor_key, url)
            
            if on_progress:
                on_progress(new_n, upd_n)

        # Flush final
        if batch:
            self.db.upsert_emails_batch(batch)

        # Efface le checkpoint (sync terminée proprement)
        self.db.set_sync_state(cursor_key, "")

        # Horodatage de la fin de sync pour ce dossier
        self.db.set_sync_state(
            f"synced_at_{folder_id}",
            datetime.now(timezone.utc).isoformat()
        )

        return new_n, upd_n

    # ── Orchestration ─────────────────────────────────────────────────────────

    def sync(
        self,
        force_full:  bool = False,
        on_status:   Callable[[str, SyncResult], None] | None = None,
    ) -> SyncResult:
        """
        Synchronise toute la boîte.
        force_full=True  → sync complète même si déjà indexé
        force_full=False → sync incrémentale depuis la dernière sync
        """
        result = SyncResult()

        if on_status:
            on_status("📁 Récupération des dossiers…", result)

        folders = self._fetch_folders()
        result.total_folders = len(folders)

        if on_status:
            on_status(f"📁 {len(folders)} dossier(s) trouvé(s) — démarrage…", result)

        for i, folder in enumerate(folders):
            fid      = folder["id"]
            fname    = folder["display_path"]
            expected = folder.get("total_item_count", 0)
            label    = f"[{i+1}/{len(folders)}] {fname}"

            # Vérifie s'il y a un checkpoint intra-dossier non terminé
            cursor = self.db.get_sync_state(f"cursor2_{fid}") or ""
            synced_at = self.db.get_sync_state(f"synced_at_{fid}") or ""

            if force_full:
                # Reset du checkpoint pour ce dossier
                self.db.set_sync_state(f"cursor2_{fid}", "")
                since_dt = None
            elif cursor:
                # Reprise d'un dossier interrompu
                since_dt = None  # on reprend depuis le checkpoint, pas de filtre date
                if on_status:
                    on_status(f"🔁 Reprise : {label}", result)
            elif synced_at and expected > 0:
                # Sync incrémentale : uniquement depuis la dernière sync
                # Marge de 10 min pour ne pas rater d'emails
                dt = datetime.fromisoformat(synced_at)
                dt = dt - timedelta(minutes=10)
                since_dt = dt.strftime("%Y-%m-%dT%H:%M:%SZ")
            else:
                since_dt = None   # Pas encore indexé → full

            mode_label = "incrémental" if (since_dt and not cursor) else "complet"
            if on_status:
                on_status(
                    f"🔄 {label} — {expected} emails attendus [{mode_label}]",
                    result,
                )

            try:
                n, u = self._sync_folder(
                    folder_id=fid,
                    folder_name=folder["name"],
                    expected=expected,
                    since_dt=since_dt,
                    on_progress=lambda nw, up, r=result, lbl=label, exp=expected: (
                        on_status(
                            f"🔄 {lbl} — {r.emails_new+nw:,}/{exp} traités",
                            r,
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
            result.folders_done   += 1

            # Vérification : compte en base vs attendu
            indexed_count = self.db.count_emails_in_folder(fid)
            info = FolderSyncInfo(
                name=fname, expected=expected,
                indexed=indexed_count, new=n, updated=u,
            )
            result.folder_details.append(info)

            gap = expected - indexed_count
            if gap > 10 and not since_dt:
                result.warnings.append(
                    f"⚠️ {fname} : {indexed_count}/{expected} emails indexés "
                    f"({gap} manquants)"
                )

            if on_status:
                check = "✅" if gap <= 10 else "⚠️"
                on_status(
                    f"{check} {fname} → {indexed_count}/{expected} indexés "
                    f"(+{n} nouveaux, ✏{u} màj)",
                    result,
                )

        self.db.set_sync_state("last_full_sync", datetime.utcnow().isoformat())
        return result