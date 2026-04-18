import requests
import sqlite3
import json
from msal import ConfidentialClientApplication
import config
import os
from datetime import datetime

class EmailFetcher:

    def __init__(self, token=None, db_path=None, user_email=None):
        """
        token: token OAuth reçu de Streamlit (optional)
        db_path: chemin vers la DB utilisateur (optional)
        user_email: email de l'utilisateur (optional)
        """
        self.token = token or self._get_token()
        self.db_path = db_path or config.DB_PATH
        self.user_email = user_email or config.USER_EMAIL
        self.graph_base = config.GRAPH_ENDPOINT

        self.headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        self._init_db()  # Initialisation de la base de données

    def _get_token(self):
        """Fallback : authentification serveur si pas de token Streamlit"""
        app = ConfidentialClientApplication(
            config.CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{config.TENANT_ID}",
            client_credential=config.CLIENT_SECRET
        )
        result = app.acquire_token_for_client(scopes=config.SCOPES)
        if "access_token" in result:
            return result["access_token"]
        raise Exception(f"Erreur auth: {result.get('error_description')}")

    def _init_db(self):
        """Crée la DB avec tables emails et sync_log, et vérifie le schéma"""
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Table principale des emails
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS emails (
                id               TEXT PRIMARY KEY,
                subject          TEXT,
                sender           TEXT,
                sender_email     TEXT,
                recipients       TEXT,
                date             TEXT,
                body             TEXT,
                body_preview     TEXT,
                folder           TEXT,
                is_read          INTEGER,
                has_attachments  INTEGER,
                user_email       TEXT,
                last_synced      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Table de suivi des synchronisations
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sync_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sync_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                emails_fetched INTEGER,
                emails_updated INTEGER,
                last_email_date TEXT
            )
        """)

        # ✅ Vérification et correction du schéma (ajout de user_email si manquant)
        cursor.execute("PRAGMA table_info(emails)")
        columns = [column[1] for column in cursor.fetchall()]

        if "user_email" not in columns:
            print("⚠️ Colonne 'user_email' manquante, ajout en cours...")
            cursor.execute("ALTER TABLE emails ADD COLUMN user_email TEXT")
            conn.commit()
            print("✅ Colonne 'user_email' ajoutée avec succès !")

        conn.commit()
        conn.close()

    def get_last_sync_time(self):
        """Retourne la date du dernier email synchronisé"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT MAX(date) FROM emails
            WHERE user_email = ?
        """, (self.user_email,))
        result = cursor.fetchone()
        conn.close()

        if result and result[0]:
            return result[0]  # Format: "2024-01-15T10:30:00Z"
        return None

    # ... (le reste de votre code existant : fetch_all_emails, search_emails, etc.)
