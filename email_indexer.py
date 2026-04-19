"""
email_indexer.py – Synchronisation des emails Office 365 via Microsoft Graph API.

Stratégie de synchronisation :
  • Première sync d'un dossier  → téléchargement complet + stockage du delta token
  • Syncs suivantes             → delta query : seuls les ajouts/modifications/suppressions
                                  depuis la dernière sync sont traités
  • Force full (reset)          → suppression des delta tokens → re-sync complète
"""

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


@dataclass
class SyncResult:
    """Résultat d'une synchronisation avec statistiques détaillées."""
    total_folders:    int = 0
    folders_full:     int = 0   # dossiers synchronisés en mode complet (1ère fois)
    folders_delta:    int = 0   # dossiers synchronisés en mode incrémental
    emails_new:       int = 0   # emails ajoutés
    emails_updated:   int = 0   # emails modifiés
    emails_deleted:   int = 0   # emails supprimés
    errors:           list[str] = field(default_factory=list)

    @property
    def emails_total(self) -> int:
        return self.emails_new + self.emails_updated


class EmailIndexer:
    """Indexe tous les emails d'une boîte Office 365 dans une base SQLite par utilisateur."""

    def __init__(self, access_token: str, user_id: str):
        self.db = Database(user_id)
        self._headers = {
            "Authorization": f"Bearer {access_token}",
            "Prefer": 'outlook.body-content-type="text"',
        }

    # ── HTTP ──────────────────────────────────────────────────────────────────

    def _get(self, url: str, params: dict | None = None) -> dict:
        for attempt in range(5):
            try:
                r = requests.get(url, headers=self._headers, params=params, timeout=30)
            except requests.RequestException as e:
                if attempt == 4:
                    raise
                time.sleep(2 ** attempt)
                continue

            if r.status_code == 429:
                time.sleep(int(r.headers.get("Retry-After", 10)))
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
        url = (
            f"{GRAPH}/me/mailFolders/{parent_id}/childFolders"
            if parent_id else f"{GRAPH}/me/mailFolders"
        )
        folders: list[dict] = []
        params = {"$top": 100}

        while url:
            data = self._get(url, params if "?" not in url else None)
            for f in data.get("value", []):
                path   = f"{parent_path}/{f['displayName']}" if parent_path else f["displayName"]
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
            url    = data.get("@odata.nextLink")
            params = None

        return folders

    # ── Parsing ───────────────────────────────────────────────────────────────

    @staticmethod
    def _parse(msg: dict, folder_id: str, folder_name: str) -> dict:
        addr       = (msg.get("from") or msg.get("sender") or {}).get("emailAddress", {})
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
            "id":                msg["id"],
            "folder_id":         folder_id,
            "folder_name":       folder_name,
            "subject":           msg.get("subject") or "(Sans objet)",
            "sender_name":       addr.get("name", ""),
            "sender_email":      addr.get("address", ""),
            "recipients":        recipients,
            "cc":                cc,
            "body_preview":      msg.get("bodyPreview", ""),
            "body":              (msg.get("body") or {}).get("content", ""),
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
        on_progress: Callable[[int, int, int], None] | None = None,
    ) -> tuple[bool, int, int, int]:
        """
        Synchronise un dossier.

        Paramètres
        ----------
        force_full  : si True, ignore le delta token et re-télécharge tout.
        on_progress : callback(new, updated, deleted) appelé après chaque batch.

        Retour
        ------
        (was_incremental, new, updated, deleted)
        """
        delta_key  = f"delta_{folder_id}"
        delta_link = self.db.get_sync_state(delta_key)

        # Détermine le mode
        if force_full or not delta_link:
            was_incremental = False
            if force_full and delta_link:
                self.db.set_sync_state(delta_key, "")   # efface l'ancien delta
            url = (
                f"{GRAPH}/me/mailFolders/{folder_id}/messages/delta"
                f"?$select={MSG_SELECT}&$top=50"
            )
        else:
            was_incremental = True
            url = delta_link

        new_count     = 0
        updated_count = 0
        deleted_count = 0
        batch: list[dict] = []

        # IDs déjà en base pour distinguer ajout vs mise à jour
        existing_ids: set[str] = set()
        if was_incremental:
            # En mode delta on ne charge pas tous les IDs (trop coûteux).
            # On détecte la mise à jour à l'upsert — on considère tout comme "new"
            # côté comptage local (le delta nous donne uniquement les changements).
            pass

        while url:
            data = self._get(url)

            for msg in data.get("value", []):
                if msg.get("@removed"):
                    self.db.delete_email(msg["id"])
                    deleted_count += 1
                    continue

                parsed = self._parse(msg, folder_id, folder_name)

                # Détecte si l'email existe déjà (mise à jour vs ajout)
                existing = self.db.get_email_detail(parsed["id"])
                if existing:
                    updated_count += 1
                else:
                    new_count += 1

                batch.append(parsed)

                if len(batch) >= 100:
                    self.db.upsert_emails_batch(batch)
                    batch = []
                    if on_progress:
                        on_progress(new_count, updated_count, deleted_count)

            new_delta = data.get("@odata.deltaLink")
            if new_delta:
                self.db.set_sync_state(delta_key, new_delta)
                url = None
            else:
                url = data.get("@odata.nextLink")

        if batch:
            self.db.upsert_emails_batch(batch)

        # Horodatage de la dernière sync du dossier
        self.db.set_sync_state(
            f"folder_synced_{folder_id}",
            datetime.utcnow().isoformat()
        )

        return was_incremental, new_count, updated_count, deleted_count

    # ── Sync complète / incrémentale ──────────────────────────────────────────

    def sync(
        self,
        force_full:  bool = False,
        on_status:   Callable[[str, SyncResult], None] | None = None,
    ) -> SyncResult:
        """
        Lance la synchronisation de toute la boîte mail.

        Paramètres
        ----------
        force_full : False  → mode incrémental (delta tokens)
                     True   → réinitialise tous les delta tokens et re-télécharge tout

        Retour
        ------
        SyncResult avec le détail des opérations effectuées.
        """
        result = SyncResult()

        if on_status:
            on_status("Récupération de la liste des dossiers…", result)

        folders = self._fetch_folders()
        result.total_folders = len(folders)

        for i, folder in enumerate(folders):
            label = f"[{i+1}/{len(folders)}] {folder['display_path']}"
            if on_status:
                on_status(f"Analyse : {label}", result)

            try:
                incremental, n, u, d = self._sync_folder(
                    folder_id=folder["id"],
                    folder_name=folder["name"],
                    force_full=force_full,
                    on_progress=lambda nw, up, dl, _lbl=label: (
                        on_status(f"Sync : {_lbl}", result) if on_status else None
                    ),
                )
            except Exception as e:
                result.errors.append(f"{folder['display_path']} : {e}")
                continue

            result.emails_new     += n
            result.emails_updated += u
            result.emails_deleted += d

            if incremental:
                result.folders_delta += 1
            else:
                result.folders_full  += 1

        self.db.set_sync_state("last_full_sync", datetime.utcnow().isoformat())
        return result

    def reset_and_full_sync(
        self,
        on_status: Callable[[str, SyncResult], None] | None = None,
    ) -> SyncResult:
        """Raccourci : réinitialise tous les delta tokens et re-sync tout."""
        return self.sync(force_full=True, on_status=on_status)