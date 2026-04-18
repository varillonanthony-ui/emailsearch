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
        """Crée la DB avec tables emails et sync_log"""
        os.makedirs("data", exist_ok=True)
        conn   = sqlite3.connect(self.db_path)
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

    def fetch_all_emails(self, incremental=True, max_emails=None):
        """
        Récupère les emails avec mode incrémental
        
        incremental=True  → ne récupère que les nouveaux depuis la dernière sync
        incremental=False → récupère TOUS les emails (lent)
        max_emails=None   → pas de limite
        """
        
        last_sync = self.get_last_sync_time() if incremental else None
        
        # ✅ Construction du filtre
        filter_str = ""
        if last_sync and incremental:
            print(f"📅 Dernière synchro : {last_sync}")
            print(f"🔍 Récupération des emails NOUVEAUX depuis {last_sync}...")
            # Format ISO 8601 pour Microsoft Graph
            filter_str = f"&$filter=receivedDateTime gt {last_sync}"
        else:
            print("🆕 Synchronisation complète : récupération de TOUS les emails...")

        # ✅ URL de base avec pagination
        url = (
            f"{config.GRAPH_ENDPOINT}/me/messages"
            f"?$top=100"
            f"&$select=id,subject,from,toRecipients,receivedDateTime"
            f",bodyPreview,body,parentFolderId,isRead,hasAttachments"
            f"&$orderby=receivedDateTime desc"
            f"{filter_str}"
        )
        
        total_fetched = 0
        total_updated = 0
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        page = 1

        while url:  # ✅ Continue tant qu'il y a un nextLink
            try:
                print(f"📄 Page {page}...")
                response = requests.get(url, headers=self.headers)

                if response.status_code != 200:
                    print(f"❌ Erreur API: {response.status_code} - {response.text}")
                    break

                data = response.json()
                emails = data.get("value", [])

                if not emails:
                    print("✅ Aucun email supplémentaire")
                    break

                print(f"📧 {len(emails)} emails dans cette page")

                # ✅ UPSERT : INSERT or REPLACE pour les doublons
                for email in emails:
                    cursor.execute("""
                        INSERT OR REPLACE INTO emails 
                        (id, subject, sender, sender_email, recipients, date, body, 
                         body_preview, folder, is_read, has_attachments, user_email, last_synced)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                    """, (
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
                    ))
                    total_fetched += 1
                    
                    # Compter les mises à jour (rowcount > 0 si UPDATE)
                    if cursor.rowcount > 0:
                        total_updated += 1

                    # ✅ OPTIONNEL : limite locale si besoin
                    if max_emails and total_fetched >= max_emails:
                        print(f"⚠️ Limite de {max_emails} atteinte")
                        conn.commit()
                        self._log_sync(total_fetched, total_updated)
                        conn.close()
                        return total_fetched

                conn.commit()
                
                # ✅ Prochaine page
                url = data.get("@odata.nextLink")
                if url:
                    page += 1
                    print(f"✅ Total: {total_fetched}, page suivante...")
                else:
                    print(f"✅ C'était la dernière page")

            except Exception as e:
                print(f"❌ Erreur lors de la récupération : {e}")
                import traceback
                traceback.print_exc()
                break

        # ✅ Log la synchronisation
        self._log_sync(total_fetched, total_updated)
        conn.close()
        
        print(f"🎉 {total_fetched} emails récupérés ({total_updated} nouveaux/mis à jour)")
        return total_fetched

    def _log_sync(self, fetched, updated):
        """Enregistre les stats de synchronisation"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Récupère la date du dernier email
        cursor.execute("SELECT MAX(date) FROM emails WHERE user_email = ?", 
                       (self.user_email,))
        last_date = cursor.fetchone()[0]
        
        cursor.execute("""
            INSERT INTO sync_log (sync_time, emails_fetched, emails_updated, last_email_date)
            VALUES (CURRENT_TIMESTAMP, ?, ?, ?)
        """, (fetched, updated, last_date))
        conn.commit()
        conn.close()
