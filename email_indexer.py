"""
email_indexer.py – Synchronisation complète des emails Office 365 via Microsoft Graph API.

Corrections v3 :
  • Suppression du N+1 (get_email_detail par email) → chargement bulk des IDs existants
  • Gestion du 410 Gone (delta token expiré) → fallback automatique en sync complète
  • Taille de page augmentée à 100 pour réduire les allers-retours API
  • Découverte exhaustive des dossiers (dossiers système + récursif)
  • Sauvegarde du delta token UNIQUEMENT à la fin (pas de perte en cas d'interruption)
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

# Dossiers système garantis — couvre Archive, Clutter, etc.
WELL_KNOWN_FOLDERS = [
    "inbox", "sentitems", "deleteditems", "drafts",
    "junkemail", "archive", "clutter", "outbox",
]


@dataclass
class SyncResult:
    total_folders:  int = 0
    folders_full:   int = 0
    folders_delta:  int = 0
    emails_new:     int = 0
    emails_updated: int = 0
    emails_deleted: int = 0
    errors:         list[str] = field(default_factory=list)

    @property
    def emails_total(self) -> int:
        return self.emails_new + self.emails_updated


class EmailIndexer:

    def __init__(self, access_token: str, user_id: str):
        self.db = Database(user_id)
        self._headers = {
            "Authorization": f"Bearer {access_token}",
            "Prefer":        'outlook.body-content-type="text"',
        }

    # ── HTTP ──────────────────────────────────────────────────────────────────

    def _get(self, url: str, params: dict | None = None) -> dict:
        """GET avec retry throttling/5xx. Lève GoneError sur 410."""
        for attempt in range(5):
            try:
                r = requests.get(url, headers=self._headers,
                                 params=params, timeout=30)
            except requests.RequestException:
                if attempt == 4:
                    raise
                time.sleep(2 ** attempt)
                continue

            if r.status_code == 410:
                raise _GoneError("Delta token expiré (410 Gone)")
            if r.status_code == 429:
                time.sleep(int(r.headers.get("Retry-After", 10)))
                continue
            if r.status_code == 401:
                raise PermissionError("Token expiré – reconnectez-vous.")
            if r.status_code >= 500:
                time.sleep(2 ** attempt)
                continue
            r.raise_for_status()
            return r.json()

        raise RuntimeError(f"Échec après 5 tentatives : {url}")

    # ── Dossiers ──────────────────────────────────────────────────────────────

    def _fetch_folders(self) -> list[dict]:
        """
        Découverte exhaustive des dossiers :
        1. Dossiers système well-known (inbox, sentitems, archive…)
        2. Tous les dossiers via /mailFolders (récursif)
        Déduplique par ID.
        """
        seen:    set[str]  = set()
        folders: list[dict] = []

        def _add(f: dict, path: str, parent_id: str | None):
            if f["id"] in seen:
                return
            seen.add(f["id"])
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

        # 1. Dossiers système
        for wk in WELL_KNOWN_FOLDERS:
            try:
                f = self._get(f"{GRAPH}/me/mailFolders/{wk}")
                _add(f, f["displayName"], None)
            except Exception:
                pass  # dossier inexistant pour cet utilisateur

        # 2. Parcours récursif complet
        def _recurse(parent_id: str | None, parent_path: str):
            url = (
                f"{GRAPH}/me/mailFolders/{parent_id}/childFolders"
                if parent_id else f"{GRAPH}/me/mailFolders"
            )
            params = {"$top": 100}
            while url:
                data = self._get(url, params if "?" not in url else None)
                for f in data.get("value", []):
                    path = f"{parent_path}/{f['displayName']}" if parent_path else f["displayName"]
                    _add(f, path, parent_id)
                    if f.get("childFolderCount", 0) > 0:
                        _recurse(f["id"], path)
                url    = data.get("@odata.nextLink")
                params = None

        _recurse(None, "")
        return folders

    # ── Parsing ───────────────────────────────────────────────────────────────

    @staticmethod
    def _parse(msg: dict, folder_id: str, folder_name: str) -> dict:
        addr = (msg.get("from") or msg.get("sender") or {}).get("emailAddress", {})
        def _addrs(key):
            return "; ".join(
                r["emailAddress"]["address"]
                for r in msg.get(key, [])
                if r.get("emailAddress", {}).get("address")
            )
        return {
            "id":                msg["id"],
            "folder_id":         folder_id,
            "folder_name":       folder_name,
            "subject":           msg.get("subject") or "(Sans objet)",
            "sender_name":       addr.get("name", ""),
            "sender_email":      addr.get("address", ""),
            "recipients":        _addrs("toRecipients"),
            "cc":                _addrs("ccRecipients"),
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
        Retourne (was_incremental, new, updated, deleted).

        Stratégie :
        - Charge en BULK les IDs déjà en base (1 requête SQL) pour
          distinguer ajout vs mise à jour sans N+1.
        - Si le delta token a expiré (410), retombe automatiquement en sync complète.
        - Sauvegarde le delta token UNIQUEMENT quand @odata.deltaLink est reçu
          (fin réelle de la sync), pas avant.
        """
        delta_key  = f"delta_{folder_id}"
        delta_link = self.db.get_sync_state(delta_key) if not force_full else None

        # Nettoie un éventuel token vide laissé par un reset précédent
        if delta_link == "":
            delta_link = None

        was_incremental = bool(delta_link)
        url = delta_link or (
            f"{GRAPH}/me/mailFolders/{folder_id}/messages/delta"
            f"?$select={MSG_SELECT}&$top=100"
        )

        # ── Chargement BULK des IDs existants (1 seule requête) ───────────────
        # Permet de distinguer new vs updated sans requête par email.
        existing_ids: set[str] = self.db.get_email_ids_for_folder(folder_id)

        new_count     = 0
        updated_count = 0
        deleted_count = 0
        batch: list[dict] = []
        final_delta: str | None = None

        try:
            while url:
                data = self._get(url)

                for msg in data.get("value", []):
                    if msg.get("@removed"):
                        self.db.delete_email(msg["id"])
                        existing_ids.discard(msg["id"])
                        deleted_count += 1
                        continue

                    parsed = self._parse(msg, folder_id, folder_name)
                    if parsed["id"] in existing_ids:
                        updated_count += 1
                    else:
                        new_count += 1
                        existing_ids.add(parsed["id"])
                    batch.append(parsed)

                    if len(batch) >= 200:
                        self.db.upsert_emails_batch(batch)
                        batch = []
                        if on_progress:
                            on_progress(new_count, updated_count, deleted_count)

                # Fin de la pagination
                new_delta = data.get("@odata.deltaLink")
                if new_delta:
                    final_delta = new_delta
                    url = None
                else:
                    url = data.get("@odata.nextLink")

        except _GoneError:
            # Delta token expiré → on recommence en sync complète pour CE dossier
            if batch:
                self.db.upsert_emails_batch(batch)
            # Réinitialise et relance en mode complet
            self.db.set_sync_state(delta_key, "")
            return self._sync_folder(folder_id, folder_name,
                                     force_full=True, on_progress=on_progress)

        # Flush du dernier batch
        if batch:
            self.db.upsert_emails_batch(batch)

        # Sauvegarde du delta token (seulement si sync complète reçue)
        if final_delta:
            self.db.set_sync_state(delta_key, final_delta)

        self.db.set_sync_state(
            f"folder_synced_{folder_id}",
            datetime.utcnow().isoformat()
        )
        return was_incremental, new_count, updated_count, deleted_count

    # ── Sync complète ─────────────────────────────────────────────────────────

    def sync(
        self,
        force_full: bool = False,
        on_status:  Callable[[str, SyncResult], None] | None = None,
    ) -> SyncResult:
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
                inc, n, u, d = self._sync_folder(
                    folder_id=folder["id"],
                    folder_name=folder["name"],
                    force_full=force_full,
                    on_progress=lambda nw, up, dl, _r=result, _lbl=label: (
                        on_status(f"Sync : {_lbl}", _r) if on_status else None
                    ),
                )
            except Exception as e:
                result.errors.append(f"{folder['display_path']} : {e}")
                continue

            result.emails_new     += n
            result.emails_updated += u
            result.emails_deleted += d
            if inc:
                result.folders_delta += 1
            else:
                result.folders_full  += 1

            if on_status:
                on_status(f"✓ {label}  (+{n} / ✏{u} / 🗑{d})", result)

        self.db.set_sync_state("last_full_sync", datetime.utcnow().isoformat())
        return result


class _GoneError(Exception):
    """Delta token expiré (HTTP 410)."""