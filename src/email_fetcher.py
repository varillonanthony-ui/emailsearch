import requests
import sqlite3
import json
from msal import ConfidentialClientApplication
import config
import os

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
        
        self.headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        self._init_db()

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
        os.makedirs("data", exist_ok=True)
        conn   = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
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
                user_email       TEXT
            )
        """)
        conn.commit()
        conn.close()

    def fetch_all_emails(self, max_emails=5000):
        url = (
            f"{config.GRAPH_ENDPOINT}/users/{self.user_email}/messages"
            f"?$top=100"
            f"&$select=id,subject,from,toRecipients,receivedDateTime"
            f",bodyPreview,body,parentFolderId,isRead,hasAttachments"
            f"&$orderby=receivedDateTime desc"
        )
        total  = 0
        conn   = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        while url and total < max_emails:
            response = requests.get(url, headers=self.headers)
            data     = response.json()
            emails   = data.get("value", [])

            for email in emails:
                cursor.execute(
                    "INSERT OR IGNORE INTO emails VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                    (
                        email.get("id"),
                        email.get("subject", ""),
                        email.get("from", {}).get("emailAddress", {}).get("name", ""),
                        email.get("from", {}).get("emailAddress", {}).get("address", ""),
                        json.dumps(email.get("toRecipients", [])),
                        email.get("receivedDateTime", ""),
                        email.get("body", {}).get("content", ""),
                        email.get("bodyPreview", ""),
                        email.get("parentFolderId", ""),
                        int(email.get("isRead", False)),
                        int(email.get("hasAttachments", False)),
                        self.user_email  # ← Ajoute l'email utilisateur
                    )
                )
                total += 1

            conn.commit()
            url = data.get("@odata.nextLink")

        conn.close()
        return total
