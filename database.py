"""
database.py – Base SQLite par utilisateur.
Chaque utilisateur possède son propre fichier .db dans /data.
"""

import sqlite3
import unicodedata
from datetime import datetime
from pathlib import Path

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)


class Database:

    def __init__(self, user_id: str):
        safe = "".join(c for c in user_id if c.isalnum() or c in "-_@.")
        self.db_path = DATA_DIR / f"{safe}.db"
        self._conn: sqlite3.Connection | None = None
        self._init_db()

    # ── Connexion ─────────────────────────────────────────────────────────────

    def _conn_get(self) -> sqlite3.Connection:
        if self._conn is None:
            self._conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
            self._conn.row_factory = sqlite3.Row
            self._conn.execute("PRAGMA journal_mode=WAL")
            self._conn.execute("PRAGMA synchronous=NORMAL")
        return self._conn

    # ── Schéma ────────────────────────────────────────────────────────────────

    def _init_db(self):
        conn = self._conn_get()
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS emails (
                id                TEXT PRIMARY KEY,
                folder_id         TEXT,
                folder_name       TEXT,
                subject           TEXT,
                sender_name       TEXT,
                sender_email      TEXT,
                recipients        TEXT,
                cc                TEXT,
                body_preview      TEXT,
                body              TEXT,
                received_datetime TEXT,
                sent_datetime     TEXT,
                has_attachments   INTEGER DEFAULT 0,
                is_read           INTEGER DEFAULT 0,
                importance        TEXT DEFAULT 'normal',
                conversation_id   TEXT,
                web_link          TEXT,
                indexed_at        TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_emails_received
                ON emails(received_datetime DESC);
            CREATE INDEX IF NOT EXISTS idx_emails_folder
                ON emails(folder_id);
            CREATE TABLE IF NOT EXISTS folders (
                id                TEXT PRIMARY KEY,
                name              TEXT,
                parent_folder_id  TEXT,
                total_item_count  INTEGER DEFAULT 0,
                unread_item_count INTEGER DEFAULT 0,
                display_path      TEXT,
                last_sync         TEXT
            );
            CREATE TABLE IF NOT EXISTS sync_state (
                key        TEXT PRIMARY KEY,
                value      TEXT,
                updated_at TEXT
            );
        """)
        conn.commit()

    # ── Emails ────────────────────────────────────────────────────────────────

    def upsert_email(self, d: dict):
        self.upsert_emails_batch([d])

    def upsert_emails_batch(self, emails: list[dict]):
        if not emails:
            return
        conn = self._conn_get()
        conn.executemany("""
            INSERT INTO emails (
                id, folder_id, folder_name, subject, sender_name, sender_email,
                recipients, cc, body_preview, body, received_datetime, sent_datetime,
                has_attachments, is_read, importance, conversation_id, web_link, indexed_at
            ) VALUES (
                :id, :folder_id, :folder_name, :subject, :sender_name, :sender_email,
                :recipients, :cc, :body_preview, :body, :received_datetime, :sent_datetime,
                :has_attachments, :is_read, :importance, :conversation_id, :web_link, :indexed_at
            )
            ON CONFLICT(id) DO UPDATE SET
                folder_id         = excluded.folder_id,
                folder_name       = excluded.folder_name,
                subject           = excluded.subject,
                sender_name       = excluded.sender_name,
                sender_email      = excluded.sender_email,
                recipients        = excluded.recipients,
                cc                = excluded.cc,
                body_preview      = excluded.body_preview,
                body              = CASE WHEN excluded.body != \'\' THEN excluded.body ELSE emails.body END,
                received_datetime = excluded.received_datetime,
                sent_datetime     = excluded.sent_datetime,
                has_attachments   = excluded.has_attachments,
                is_read           = excluded.is_read,
                importance        = excluded.importance,
                conversation_id   = excluded.conversation_id,
                web_link          = excluded.web_link,
                indexed_at        = excluded.indexed_at
        """, emails)
        conn.commit()

    def delete_email(self, email_id: str):
        conn = self._conn_get()
        conn.execute("DELETE FROM emails WHERE id = ?", (email_id,))
        conn.commit()

    def get_email_detail(self, email_id: str) -> dict | None:
        cur = self._conn_get().cursor()
        cur.execute("SELECT * FROM emails WHERE id = ?", (email_id,))
        row = cur.fetchone()
        return dict(row) if row else None

    def get_email_ids_for_folder(self, folder_id: str) -> set[str]:
        cur = self._conn_get().cursor()
        cur.execute("SELECT id FROM emails WHERE folder_id = ?", (folder_id,))
        return {row["id"] for row in cur.fetchall()}

    def count_emails_in_folder(self, folder_id: str) -> int:
        cur = self._conn_get().cursor()
        cur.execute("SELECT COUNT(*) AS n FROM emails WHERE folder_id = ?", (folder_id,))
        return cur.fetchone()["n"]

    def reset_all(self):
        conn = self._conn_get()
        conn.executescript("DELETE FROM emails; DELETE FROM folders; DELETE FROM sync_state;")
        conn.commit()

    # ── Dossiers ──────────────────────────────────────────────────────────────

    def upsert_folder(self, d: dict):
        self._conn_get().execute("""
            INSERT INTO folders (id, name, parent_folder_id, total_item_count,
                unread_item_count, display_path, last_sync)
            VALUES (:id, :name, :parent_folder_id, :total_item_count,
                :unread_item_count, :display_path, :last_sync)
            ON CONFLICT(id) DO UPDATE SET
                name              = excluded.name,
                total_item_count  = excluded.total_item_count,
                unread_item_count = excluded.unread_item_count,
                display_path      = excluded.display_path
        """, d)
        self._conn_get().commit()

    def get_folders(self) -> list[dict]:
        cur = self._conn_get().cursor()
        cur.execute("SELECT * FROM folders ORDER BY display_path")
        return [dict(r) for r in cur.fetchall()]

    # ── Sync state ────────────────────────────────────────────────────────────

    def get_sync_state(self, key: str) -> str | None:
        cur = self._conn_get().cursor()
        cur.execute("SELECT value FROM sync_state WHERE key = ?", (key,))
        row = cur.fetchone()
        return row["value"] if row else None

    def set_sync_state(self, key: str, value: str):
        self._conn_get().execute("""
            INSERT INTO sync_state (key, value, updated_at) VALUES (?, ?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at
        """, (key, value, datetime.utcnow().isoformat()))
        self._conn_get().commit()

    # ── Recherche ─────────────────────────────────────────────────────────────

    EXCLUDED_FOLDERS = (
        "Éléments envoyés", "Sent Items", "Sent",
        "Éléments supprimés", "Deleted Items", "Trash", "Corbeille",
        "Courrier indésirable", "Junk Email", "Spam",
        "Brouillons", "Drafts",
    )

    @staticmethod
    def _normalize(text: str) -> str:
        """Normalisation Unicode : supprime accents, met en minuscules."""
        t = unicodedata.normalize("NFD", text.lower())
        return "".join(c for c in t if unicodedata.category(c) != "Mn")

    def search_emails(
        self,
        keywords:      list[str],
        folder_filter: str = "all",
        folder_ids:    list[str] | None = None,
        date_from:     str | None = None,
        date_to:       str | None = None,
        limit:         int = 25,
        offset:        int = 0,
    ) -> tuple[list[dict], int]:
        """
        Recherche multi-mots-clés (logique ET) avec filtre date/dossier.
        Étape 1 — SQL : filtre dossier + date (réduit le dataset).
        Étape 2 — Python : normalisation Unicode + logique ET sur tous les mots-clés.
        """
        clean_kws = [self._normalize(kw) for kw in keywords if kw.strip()]
        if not clean_kws:
            return [], 0

        conds:  list[str] = []
        params: list      = []

        if folder_filter == "no_sent_deleted":
            ph = ",".join("?" * len(self.EXCLUDED_FOLDERS))
            conds.append(f"folder_name NOT IN ({ph})")
            params.extend(self.EXCLUDED_FOLDERS)
        elif folder_filter == "specific" and folder_ids:
            ph = ",".join("?" * len(folder_ids))
            conds.append(f"folder_id IN ({ph})")
            params.extend(folder_ids)

        if date_from:
            conds.append("received_datetime >= ?")
            params.append(date_from)
        if date_to:
            conds.append("received_datetime < ?")
            params.append(date_to + "T23:59:59")

        where = ("WHERE " + " AND ".join(conds)) if conds else ""

        cur = self._conn_get().cursor()
        cur.execute(f"""
            SELECT id, folder_id, folder_name, subject,
                   sender_name, sender_email, recipients,
                   body_preview, body,
                   received_datetime, sent_datetime,
                   has_attachments, is_read, importance,
                   web_link, conversation_id
            FROM emails {where}
            ORDER BY received_datetime DESC
        """, params)

        matched = [
            dict(row) for row in cur.fetchall()
            if all(kw in self._normalize(" ".join([
                row["subject"]      or "",
                row["sender_name"]  or "",
                row["sender_email"] or "",
                row["recipients"]   or "",
                row["body_preview"] or "",
                row["body"]         or "",
            ])) for kw in clean_kws)
        ]

        return matched[offset: offset + limit], len(matched)

    # ── Statistiques ──────────────────────────────────────────────────────────

    def get_stats(self) -> dict:
        cur = self._conn_get().cursor()
        cur.execute("SELECT COUNT(*) AS n FROM emails")
        total_emails = cur.fetchone()["n"]
        cur.execute("SELECT COUNT(*) AS n FROM folders")
        total_folders = cur.fetchone()["n"]
        cur.execute("""
            SELECT folder_name, COUNT(*) AS cnt FROM emails
            GROUP BY folder_name ORDER BY cnt DESC LIMIT 20
        """)
        by_folder = [dict(r) for r in cur.fetchall()]
        return {
            "total_emails":  total_emails,
            "total_folders": total_folders,
            "by_folder":     by_folder,
            "last_sync":     self.get_sync_state("last_full_sync"),
        }

    def close(self):
        if self._conn:
            self._conn.close()
            self._conn = None