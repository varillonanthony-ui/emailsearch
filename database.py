"""
database.py – Base SQLite par utilisateur avec FTS5 pour la recherche plein texte.
Chaque utilisateur possède son propre fichier .db dans le dossier /data.
"""

import sqlite3
import os
from datetime import datetime
from pathlib import Path

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)


class Database:
    """Encapsule toutes les opérations SQLite pour un utilisateur donné."""

    def __init__(self, user_id: str):
        # Sécurise le nom de fichier
        safe = "".join(c for c in user_id if c.isalnum() or c in "-_@.")
        self.db_path = DATA_DIR / f"{safe}.db"
        self._conn: sqlite3.Connection | None = None
        self._init_db()

    # ── Connexion ──────────────────────────────────────────────────────────────

    def _conn_get(self) -> sqlite3.Connection:
        if self._conn is None:
            self._conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
            self._conn.row_factory = sqlite3.Row
            self._conn.execute("PRAGMA journal_mode=WAL")
            self._conn.execute("PRAGMA synchronous=NORMAL")
        return self._conn

    # ── Initialisation schéma ──────────────────────────────────────────────────

    def _init_db(self):
        conn = self._conn_get()
        cur = conn.cursor()

        # Table principale des emails
        cur.execute("""
            CREATE TABLE IF NOT EXISTS emails (
                id                  TEXT PRIMARY KEY,
                folder_id           TEXT,
                folder_name         TEXT,
                subject             TEXT,
                sender_name         TEXT,
                sender_email        TEXT,
                recipients          TEXT,
                cc                  TEXT,
                body_preview        TEXT,
                body                TEXT,
                received_datetime   TEXT,
                sent_datetime       TEXT,
                has_attachments     INTEGER DEFAULT 0,
                is_read             INTEGER DEFAULT 0,
                importance          TEXT DEFAULT 'normal',
                conversation_id     TEXT,
                web_link            TEXT,
                indexed_at          TEXT
            )
        """)

        # Index pour les tris courants
        cur.execute("CREATE INDEX IF NOT EXISTS idx_emails_received ON emails(received_datetime DESC)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_emails_folder   ON emails(folder_id)")

        # Table FTS5 autonome (copie des champs cherchables)
        cur.execute("""
            CREATE VIRTUAL TABLE IF NOT EXISTS emails_fts
            USING fts5(
                id,
                subject,
                sender_name,
                sender_email,
                recipients,
                body_preview,
                body,
                folder_name,
                tokenize = 'unicode61 remove_diacritics 1'
            )
        """)

        # Table des dossiers
        cur.execute("""
            CREATE TABLE IF NOT EXISTS folders (
                id                  TEXT PRIMARY KEY,
                name                TEXT,
                parent_folder_id    TEXT,
                total_item_count    INTEGER DEFAULT 0,
                unread_item_count   INTEGER DEFAULT 0,
                display_path        TEXT,
                last_sync           TEXT
            )
        """)

        # État de synchronisation (delta links, timestamps…)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS sync_state (
                key         TEXT PRIMARY KEY,
                value       TEXT,
                updated_at  TEXT
            )
        """)

        conn.commit()

    # ── Emails ────────────────────────────────────────────────────────────────

    def upsert_email(self, d: dict):
        """Insère ou met à jour un email (et son entrée FTS)."""
        self.upsert_emails_batch([d])

    def upsert_emails_batch(self, emails: list[dict]):
        """Insère / met à jour un lot d'emails de façon atomique."""
        if not emails:
            return
        conn = self._conn_get()
        cur = conn.cursor()

        for d in emails:
            # Upsert table principale
            cur.execute("""
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
                    body              = excluded.body,
                    received_datetime = excluded.received_datetime,
                    sent_datetime     = excluded.sent_datetime,
                    has_attachments   = excluded.has_attachments,
                    is_read           = excluded.is_read,
                    importance        = excluded.importance,
                    conversation_id   = excluded.conversation_id,
                    web_link          = excluded.web_link,
                    indexed_at        = excluded.indexed_at
            """, d)

            # Synchronisation FTS5 : supprime l'ancienne entrée, réinsère
            cur.execute("DELETE FROM emails_fts WHERE id = ?", (d["id"],))
            cur.execute("""
                INSERT INTO emails_fts (id, subject, sender_name, sender_email,
                    recipients, body_preview, body, folder_name)
                VALUES (:id, :subject, :sender_name, :sender_email,
                    :recipients, :body_preview, :body, :folder_name)
            """, d)

        conn.commit()

    def delete_email(self, email_id: str):
        """Supprime un email et son entrée FTS."""
        conn = self._conn_get()
        cur = conn.cursor()
        cur.execute("DELETE FROM emails     WHERE id = ?", (email_id,))
        cur.execute("DELETE FROM emails_fts WHERE id = ?", (email_id,))
        conn.commit()

    def get_email_ids_for_folder(self, folder_id: str) -> set[str]:
        """
        Retourne l'ensemble des IDs d'emails déjà indexés pour un dossier.
        Utilisé par l'indexeur pour distinguer ajout vs mise à jour en 1 requête SQL
        (évite le N+1 : pas de get_email_detail() par email dans la boucle de sync).
        """
        cur = self._conn_get().cursor()
        cur.execute("SELECT id FROM emails WHERE folder_id = ?", (folder_id,))
        return {row["id"] for row in cur.fetchall()}

    def count_emails_in_folder(self, folder_id: str) -> int:
        """Compte les emails indexés pour un dossier (pour vérification post-sync)."""
        cur = self._conn_get().cursor()
        cur.execute("SELECT COUNT(*) AS n FROM emails WHERE folder_id = ?", (folder_id,))
        return cur.fetchone()["n"]

    def upsert_email(self, d: dict):
        """Alias pour mise à jour d'un email unique (ex: mise en cache du corps)."""
        self.upsert_emails_batch([d])

    def reset_all(self):
        """
        Supprime toutes les données (emails, FTS, dossiers, états de sync).
        Utilisé pour repartir d'une base propre.
        """
        conn = self._conn_get()
        conn.execute("DELETE FROM emails")
        conn.execute("DELETE FROM emails_fts")
        conn.execute("DELETE FROM folders")
        conn.execute("DELETE FROM sync_state")
        conn.commit()

    def get_email_detail(self, email_id: str) -> dict | None:
        """Retourne tous les champs d'un email (dont le corps complet)."""
        cur = self._conn_get().cursor()
        cur.execute("SELECT * FROM emails WHERE id = ?", (email_id,))
        row = cur.fetchone()
        return dict(row) if row else None

    # ── Dossiers ──────────────────────────────────────────────────────────────

    def upsert_folder(self, d: dict):
        conn = self._conn_get()
        conn.execute("""
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
        conn.commit()

    def get_folders(self) -> list[dict]:
        cur = self._conn_get().cursor()
        cur.execute("SELECT * FROM folders ORDER BY display_path")
        return [dict(r) for r in cur.fetchall()]

    # ── État de sync ──────────────────────────────────────────────────────────

    def get_sync_state(self, key: str) -> str | None:
        cur = self._conn_get().cursor()
        cur.execute("SELECT value FROM sync_state WHERE key = ?", (key,))
        row = cur.fetchone()
        return row["value"] if row else None

    def set_sync_state(self, key: str, value: str):
        self._conn_get().execute("""
            INSERT INTO sync_state (key, value, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at
        """, (key, value, datetime.utcnow().isoformat()))
        self._conn_get().commit()

    # ── Recherche ─────────────────────────────────────────────────────────────

    # Dossiers système exclus du filtre "hors envoyés/supprimés"
    EXCLUDED_FOLDER_NAMES = (
        "Éléments envoyés", "Sent Items", "Sent",
        "Éléments supprimés", "Deleted Items", "Trash", "Corbeille",
        "Courrier indésirable", "Junk Email", "Spam",
        "Brouillons", "Drafts",
    )

    # Colonnes de recherche (concaténées pour une comparaison unique)
    SEARCH_COLS = ("subject", "sender_name", "sender_email",
                   "recipients", "body_preview", "body")

    @staticmethod
    def _normalize(text: str) -> str:
        """
        Normalise un texte pour la comparaison insensible à la casse et aux accents.
        Utilise unicodedata pour gérer correctement les accents français (é→e, è→e…).
        """
        import unicodedata
        text = unicodedata.normalize("NFD", text.lower())
        return "".join(c for c in text if unicodedata.category(c) != "Mn")

    def search_emails(
        self,
        keywords: list[str],
        folder_filter: str = "all",
        folder_ids: list[str] | None = None,
        date_from: str | None = None,   # "YYYY-MM-DD" ou None
        date_to:   str | None = None,   # "YYYY-MM-DD" ou None
        limit: int = 20,
        offset: int = 0,
    ) -> tuple[list[dict], int]:
        """
        Recherche multi-mots-clés avec logique ET, filtre date et normalisation Unicode.

        Stratégie :
        1. SQL : filtre dossier + date (réduit le dataset côté base)
        2. Python : filtre ET strict sur les mots-clés avec normalisation Unicode
           (accents, casse) — 100% fiable.

        Paramètres date : "YYYY-MM-DD" ou None (= pas de filtre).
        """
        conn = self._conn_get()
        cur  = conn.cursor()

        # Nettoyage et normalisation des mots-clés
        clean_kws = [self._normalize(kw) for kw in keywords if kw.strip()]
        if not clean_kws:
            return [], 0

        # ── Filtres SQL (dossier + date) ─────────────────────────────────────
        sql_conditions: list[str] = []
        sql_params:     list      = []

        if folder_filter == "no_sent_deleted":
            ph = ",".join("?" * len(self.EXCLUDED_FOLDER_NAMES))
            sql_conditions.append(f"e.folder_name NOT IN ({ph})")
            sql_params.extend(self.EXCLUDED_FOLDER_NAMES)
        elif folder_filter == "specific" and folder_ids:
            ph = ",".join("?" * len(folder_ids))
            sql_conditions.append(f"e.folder_id IN ({ph})")
            sql_params.extend(folder_ids)

        if date_from:
            sql_conditions.append("e.received_datetime >= ?")
            sql_params.append(date_from)          # ISO compare works lexicographically

        if date_to:
            # Inclut toute la journée de date_to
            sql_conditions.append("e.received_datetime < ?")
            dt_end = date_to + "T23:59:59"
            sql_params.append(dt_end)

        where = ("WHERE " + " AND ".join(sql_conditions)) if sql_conditions else ""

        # ── Récupération des emails du périmètre ─────────────────────────────
        cur.execute(f"""
            SELECT e.id, e.folder_id, e.folder_name, e.subject,
                   e.sender_name, e.sender_email, e.recipients,
                   e.body_preview, e.body,
                   e.received_datetime, e.sent_datetime,
                   e.has_attachments, e.is_read, e.importance,
                   e.web_link, e.conversation_id
            FROM emails e
            {where}
            ORDER BY e.received_datetime DESC
        """, sql_params)

        rows = cur.fetchall()

        # ── Filtrage Python multi-mots-clés (normalisation Unicode) ──────────
        matched: list[dict] = []
        for row in rows:
            haystack = self._normalize(" ".join([
                row["subject"]      or "",
                row["sender_name"]  or "",
                row["sender_email"] or "",
                row["recipients"]   or "",
                row["body_preview"] or "",
                row["body"]         or "",
            ]))
            if all(kw in haystack for kw in clean_kws):
                matched.append(dict(row))

        total   = len(matched)
        results = matched[offset: offset + limit]
        return results, total

    # ── Statistiques ──────────────────────────────────────────────────────────

    def get_stats(self) -> dict:
        cur = self._conn_get().cursor()
        cur.execute("SELECT COUNT(*) AS n FROM emails")
        total_emails = cur.fetchone()["n"]
        cur.execute("SELECT COUNT(*) AS n FROM folders")
        total_folders = cur.fetchone()["n"]
        cur.execute("""
            SELECT folder_name, COUNT(*) AS cnt
            FROM emails
            GROUP BY folder_name
            ORDER BY cnt DESC
            LIMIT 20
        """)
        by_folder = [dict(r) for r in cur.fetchall()]
        last_sync = self.get_sync_state("last_full_sync")
        return {
            "total_emails": total_emails,
            "total_folders": total_folders,
            "by_folder": by_folder,
            "last_sync": last_sync,
        }

    def close(self):
        if self._conn:
            self._conn.close()
            self._conn = None