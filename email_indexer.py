"""
email_indexer.py – Synchronisation des emails depuis Office 365 via Microsoft Graph API.
Utilise les delta tokens pour n'indexer que les nouveaux messages lors des syncs suivantes.
"""

import time
from datetime import datetime
from typing import Callable
import requests

from database import Database

GRAPH = "https://graph.microsoft.com/v1.0"

# Sélection des champs à récupérer (économise la bande passante)
MSG_SELECT = (
    "id,subject,from,sender,toRecipients,ccRecipients,"
    "bodyPreview,body,receivedDateTime,sentDateTime,"
    "hasAttachments,isRead,importance,conversationId,webLink"
)


class EmailIndexer:
    """Indexe tous les emails d'une boîte Office 365 dans une base SQLite."""

    def __init__(self, access_token: str, user_id: str):
        self.db = Database(user_id)
        self._headers = {
            "Authorization": f"Bearer {access_token}",
            # Demande le corps en texte brut (moins volumineux que HTML)
            "Prefer": 'outlook.body-content-type="text"',
        }

    # ── HTTP ──────────────────────────────────────────────────────────────────

    def _get(self, url: str, params: dict | None = None) -> dict:
        """GET avec retry sur throttling (429) et erreurs transitoires."""
        for attempt in range(5):
            try:
                r = requests.get(url, headers=self._headers, params=params, timeout=30)
            except requests.RequestException as e:
                if attempt == 4:
                    raise
                time.sleep(2 ** attempt)
                continue

            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", 10))
                time.sleep(wait)
                continue
            if r.status_code == 401:
                raise PermissionError("Token expiré ou invalide – reconnectez-vous.")
            if r.status_code >= 500:
                time.sleep(2 ** attempt)
                continue
            r.raise_for_status()
            return r.json()

        raise RuntimeError(f"Impossible d'accéder à {url} après 5 tentatives.")

    # ── Dossiers ──────────────────────────────────────────────────────────────

    def _fetch_folders(self, parent_id: str | None = None, parent_path: str = "") -> list[dict]:
        """Récupère récursivement tous les dossiers de la boîte mail."""
        url = (
            f"{GRAPH}/me/mailFolders/{parent_id}/childFolders"
            if parent_id
            else f"{GRAPH}/me/mailFolders"
        )
        folders: list[dict] = []
        params = {"$top": 100}

        while url:
            data = self._get(url, params if "?" not in url else None)
            for f in data.get("value", []):
                path = f"{parent_path}/{f['displayName']}" if parent_path else f["displayName"]
                record = {
                    "id":                f["id"],
                    "name":              f["displayName"],
                    "parent_folder_id":  parent_id or "",
                    "total_item_count":  f.get("totalItemCount", 0),
                    "unread_item_count": f.get("unreadItemCount", 0),
                    "display_path":      path,
                    "last_sync":         None,
                }
                folders.append(record)
                self.db.upsert_folder(record)

                if f.get("childFolderCount", 0) > 0:
                    folders.extend(self._fetch_folders(f["id"], path))

            url = data.get("@odata.nextLink")
            params = None

        return folders

    # ── Parsing email ─────────────────────────────────────────────────────────

    @staticmethod
    def _parse(msg: dict, folder_id: str, folder_name: str) -> dict:
        """Convertit un objet Graph API en dictionnaire pour la DB."""
        addr = (msg.get("from") or msg.get("sender") or {}).get("emailAddress", {})
        recipients = "; ".join(
            r["emailAddress"]["address"]
            for r in msg.get("toRecipients", [])
            if r.get("emailAddress", {}).get("address")
        )
        cc = "; ".join(
            r["emailAddress"]["address"]
            for r in msg.get("ccRecipients", [])
            if r.get("emailAddress", {}).get("address")
        )
        return {
            "id":                 msg["id"],
            "folder_id":          folder_id,
            "folder_name":        folder_name,
            "subject":            msg.get("subject") or "(Sans objet)",
            "sender_name":        addr.get("name", ""),
            "sender_email":       addr.get("address", ""),
            "recipients":         recipients,
            "cc":                 cc,
            "body_preview":       msg.get("bodyPreview", ""),
            "body":               (msg.get("body") or {}).get("content", ""),
            "received_datetime":  msg.get("receivedDateTime", ""),
            "sent_datetime":      msg.get("sentDateTime", ""),
            "has_attachments":    1 if msg.get("hasAttachments") else 0,
            "is_read":            1 if msg.get("isRead") else 0,
            "importance":         msg.get("importance", "normal"),
            "conversation_id":    msg.get("conversationId", ""),
            "web_link":           msg.get("webLink", ""),
            "indexed_at":         datetime.utcnow().isoformat(),
        }

    # ── Sync dossier ──────────────────────────────────────────────────────────

    def _sync_folder(
        self,
        folder_id: str,
        folder_name: str,
        on_progress: Callable[[int], None] | None = None,
    ) -> int:
        """
        Synchronise un dossier via l'API delta Graph.
        Retourne le nombre d'emails traités.
        """
        delta_key  = f"delta_{folder_id}"
        delta_link = self.db.get_sync_state(delta_key)

        if delta_link:
            # Sync incrémentale : seuls les changements depuis le dernier delta
            url = delta_link
        else:
            # Sync complète
            url = (
                f"{GRAPH}/me/mailFolders/{folder_id}/messages/delta"
                f"?$select={MSG_SELECT}&$top=50"
            )

        total = 0
        batch: list[dict] = []

        while url:
            data = self._get(url)

            for msg in data.get("value", []):
                # Message supprimé détecté lors d'un delta
                if msg.get("@removed"):
                    self.db.delete_email(msg["id"])
                    continue

                batch.append(self._parse(msg, folder_id, folder_name))
                total += 1

                if len(batch) >= 100:
                    self.db.upsert_emails_batch(batch)
                    batch = []
                    if on_progress:
                        on_progress(total)

            new_delta = data.get("@odata.deltaLink")
            if new_delta:
                self.db.set_sync_state(delta_key, new_delta)
                url = None
            else:
                url = data.get("@odata.nextLink")

        if batch:
            self.db.upsert_emails_batch(batch)

        return total

    # ── Sync complète ─────────────────────────────────────────────────────────

    def full_sync(self, on_status: Callable[[str, int], None] | None = None) -> int:
        """
        Synchronise l'intégralité de la boîte mail.
        on_status(message, total_indexé) est appelé régulièrement.
        Retourne le nombre total d'emails indexés/mis à jour.
        """
        if on_status:
            on_status("Récupération de la liste des dossiers…", 0)

        folders = self._fetch_folders()
        grand_total = 0

        for i, folder in enumerate(folders):
            label = f"[{i+1}/{len(folders)}] {folder['display_path']}"
            if on_status:
                on_status(f"Sync : {label}", grand_total)

            def _cb(count, _label=label, _gt=grand_total):
                if on_status:
                    on_status(f"Sync : {_label}", _gt + count)

            n = self._sync_folder(folder["id"], folder["name"], on_progress=_cb)
            grand_total += n

        self.db.set_sync_state("last_full_sync", datetime.utcnow().isoformat())
        return grand_total