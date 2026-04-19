"""
email_indexer.py – Synchronisation complète des emails Office 365.

Corrections v4 :
  • Chargement bulk des IDs (1 requête SQL par dossier, pas N+1)
  • Gestion 410 Gone (delta expiré) → fallback full automatique
  • Sauvegarde du delta token UNIQUEMENT à la fin complète
  • Découverte dossiers : well-known + récursif complet + déduplication
  • Sync par checkpoint : chaque dossier terminé est marqué
    → si la session est interrompue, les dossiers déjà faits sont sautés
  • Body text nettoyé avant insertion (évite les \x00 qui cassent SQLite)
  • Page size 100, timeout 60s
"""

import re
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable
import requests

from database import Database

GRAPH      = "https://graph.microsoft.com/v1.0"
MSG_SELECT = (
    "id,subject,from,sender,toRecipients,ccRecipients,"
    "bodyPreview,body,receivedDateTime,sentDateTime,"
    "hasAttachments,isRead,importance,conversationId,webLink"
)

# Dossiers système explicites à toujours inclure
WELL_KNOWN = [
    "inbox", "sentitems", "deleteditems", "drafts",
    "junkemail", "archive", "clutter", "outbox",
]


@dataclass
class SyncResult:
    total_folders:  int = 0
    folders_done:   int = 0   # terminés cette session
    folders_skip:   int = 0   # déjà faits (checkpoint)
    emails_new:     int = 0
    emails_updated: int = 0
    emails_deleted: int = 0
    errors:         list[str] = field(default_factory=list)

    @property
    def emails_total(self) -> int:
        return self.emails_new + self.emails_updated


def _clean(text: str | None) -> str:
    """Supprime les caractères nuls et normalise les espaces."""
    if not text:
        return ""
    text = text.replace("\x00", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


class _GoneError(Exception):
    """Delta token expiré (HTTP 410)."""


class EmailIndexer:

    def __init__(self, access_token: str, user_id: str):
        self.db = Database(user_id)
        self._tok = access_token
        self._headers = {
            "Authorization": f"Bearer {access_token}",
            "Prefer":        'outlook.body-content-type="text"',
        }

    # ── HTTP ──────────────────────────────────────────────────────────────────

    def _get(self, url: str, params: dict | None = None) -> dict:
        for attempt in range(6):
            try:
                r = requests.get(url, headers=self._headers,
                                 params=params, timeout=60)
            except requests.RequestException as exc:
                if attempt == 5:
                    raise RuntimeError(f"Réseau : {exc}") from exc
                time.sleep(2 ** attempt)
                continue

            if r.status_code == 410:
                raise _GoneError("Delta token expiré (410)")
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", 15))
                time.sleep(wait)
                continue
            if r.status_code == 401:
                raise PermissionError("Token Microsoft expiré – reconnectez-vous.")
            if r.status_code in (500, 502, 503, 504):
                time.sleep(2 ** attempt)
                continue
            try:
                r.raise_for_status()
            except Exception:
                raise RuntimeError(f"HTTP {r.status_code} sur {url}: {r.text[:200]}")
            return r.json()
        raise RuntimeError(f"Abandon après 6 tentatives : {url}")

    # ── Dossiers ──────────────────────────────────────────────────────────────

    def _fetch_folders(self) -> list[dict]:
        """
        Découverte exhaustive :
        1. Dossiers well-known (inbox, sentitems, archive…)
        2. Parcours récursif /mailFolders
        Déduplique par ID.
        """
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

        # 1. Well-known
        for wk in WELL_KNOWN:
            try:
                f = self._get(f"{GRAPH}/me/mailFolders/{wk}")
                _save(f, f["displayName"], None)
            except Exception:
                pass

        # 2. Récursif
        def _recurse(parent_id: str | None, parent_path: str):
            url = (
                f"{GRAPH}/me/mailFolders/{parent_id}/childFolders"
                if parent_id else f"{GRAPH}/me/mailFolders"
            )
            params = {"$top": 100}
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

        body = _clean((msg.get("body") or {}).get("content", ""))

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
            "body":              body,
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
        folder_id:   str,
        folder_name: str,
        force_full:  bool = False,
        on_progress: Callable | None = None,
    ) -> tuple[bool, int, int, int]:
        """
        Retourne (was_incremental, new, updated, deleted).

        • Chargement bulk des IDs existants (1 SELECT, pas N+1)
        • Fallback full automatique si 410
        • Delta token sauvegardé seulement à la FIN complète
        """
        delta_key  = f"delta_{folder_id}"
        delta_link = None if force_full else (self.db.get_sync_state(delta_key) or None)

        was_inc = bool(delta_link)
        url = delta_link or (
            f"{GRAPH}/me/mailFolders/{folder_id}/messages/delta"
            f"?$select={MSG_SELECT}&$top=100"
        )

        # Bulk load des IDs déjà en base (1 seule requête SQL)
        existing: set[str] = self.db.get_email_ids_for_folder(folder_id)

        new_n = upd_n = del_n = 0
        batch: list[dict] = []
        final_delta: str | None = None

        try:
            while url:
                data = self._get(url)

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

                    if len(batch) >= 200:
                        self.db.upsert_emails_batch(batch)
                        batch = []
                        if on_progress:
                            on_progress(new_n, upd_n, del_n)

                if data.get("@odata.deltaLink"):
                    final_delta = data["@odata.deltaLink"]
                    url = None
                else:
                    url = data.get("@odata.nextLink")

        except _GoneError:
            # Delta expiré → on vide et recommence en full pour ce dossier
            if batch:
                self.db.upsert_emails_batch(batch)
            self.db.set_sync_state(delta_key, "")
            return self._sync_folder(folder_id, folder_name,
                                     force_full=True, on_progress=on_progress)

        if batch:
            self.db.upsert_emails_batch(batch)

        # Sauvegarde delta seulement si pagination complète reçue
        if final_delta:
            self.db.set_sync_state(delta_key, final_delta)

        self.db.set_sync_state(f"folder_done_{folder_id}", datetime.utcnow().isoformat())
        return was_inc, new_n, upd_n, del_n

    # ── Orchestration ─────────────────────────────────────────────────────────

    def sync(
        self,
        force_full: bool = False,
        on_status:  Callable[[str, SyncResult], None] | None = None,
    ) -> SyncResult:
        """
        Synchronise toute la boîte.

        Checkpoint par dossier : si la session est interrompue et relancée
        en mode incrémental, les dossiers déjà terminés sont sautés
        (leur delta token les mettra à jour très vite).
        """
        result = SyncResult()

        if on_status:
            on_status("📁 Récupération des dossiers…", result)

        folders = self._fetch_folders()
        result.total_folders = len(folders)

        for i, folder in enumerate(folders):
            fid   = folder["id"]
            fname = folder["display_path"]
            label = f"[{i+1}/{len(folders)}] {fname}"

            # Checkpoint : dossier déjà synchronisé cette session ?
            already_done = (
                not force_full
                and bool(self.db.get_sync_state(f"folder_done_{fid}"))
                and bool(self.db.get_sync_state(f"delta_{fid}"))
            )
            if already_done:
                result.folders_skip += 1
                if on_status:
                    on_status(f"⏭ Déjà indexé : {fname}", result)
                continue

            if on_status:
                on_status(f"🔄 {label}", result)

            try:
                inc, n, u, d = self._sync_folder(
                    folder_id=fid,
                    folder_name=folder["name"],
                    force_full=force_full,
                    on_progress=lambda nw, up, dl: (
                        on_status(f"🔄 {label} (+{nw}✉)", result) if on_status else None
                    ),
                )
            except PermissionError:
                raise
            except Exception as exc:
                result.errors.append(f"{fname} : {exc}")
                if on_status:
                    on_status(f"⚠️ Erreur sur {fname} : {exc}", result)
                continue

            result.emails_new     += n
            result.emails_updated += u
            result.emails_deleted += d
            result.folders_done   += 1

            if on_status:
                mode = "delta" if inc else "complet"
                on_status(
                    f"✅ {fname} ({mode}) — +{n} nouveaux / {u} màj / {d} supprimés",
                    result,
                )

        self.db.set_sync_state("last_full_sync", datetime.utcnow().isoformat())
        return result