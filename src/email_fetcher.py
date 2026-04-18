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

    def fetch_all_emails(self, max_emails=None):
        """
        max_emails: limite maximale (None = pas de limite)
        """
        # ✅ CORRECTION : pas de limite par défaut
        url = (
            f"{config.GRAPH_ENDPOINT}/me/messages"
            f"?$top=100"
            f"&$select=id,subject,from,toRecipients,receivedDateTime"
            f",bodyPreview,body,parentFolderId,isRead,hasAttachments"
            f"&$orderby=receivedDateTime desc"
        )
        total  = 0
        conn   = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        while url:  # ✅ Continue tant qu'il y a un nextLink
            print(f"🔄 Fetch page: {url[:80]}...")
            response = requests.get(url, headers=self.headers)
            
            if response.status_code != 200:
                print(f"❌ Erreur API: {response.status_code} - {response.text}")
                break
            
            data   = response.json()
            emails = data.get("value", [])
            
            print(f"📧 {len(emails)} emails dans cette page")

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
                        self.user_email
                    )
                )
                total += 1

                # ✅ OPTIONNEL : limite locale si besoin
                if max_emails and total >= max_emails:
                    print(f"⚠️ Limite de {max_emails} atteinte, arrêt")
                    conn.commit()
                    conn.close()
                    return total

            conn.commit()
            url = data.get("@odata.nextLink")
            print(f"✅ Total: {total}, NextLink: {bool(url)}")

        conn.close()
        print(f"🎉 {total} emails sauvegardés !")
        return total
