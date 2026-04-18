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

    def _log_sync(self, fetched, updated):
        """Enregistre une synchronisation dans le log"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute("""
            INSERT INTO sync_log (emails_fetched, emails_updated, last_email_date)
            VALUES (?, ?, ?)
        """, (fetched, updated, datetime.now().isoformat()))

        conn.commit()
        conn.close()

    def fetch_all_emails(self, incremental=True, max_emails=None):
        """
        Récupère les emails depuis Microsoft Graph

        Args:
            incremental (bool): Si True, ne récupère que les nouveaux emails depuis la dernière synchronisation
            max_emails (int): Limite le nombre d'emails à récupérer (None = pas de limite)

        Returns:
            int: Nombre total d'emails récupérés
        """
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Initialisation des compteurs
        total_fetched = 0
        total_updated = 0
        page = 1
        next_link = None

        # Construction de l'URL de l'API Graph
        url = f"{self.graph_base}/me/mailFolders/inbox/messages"
        params = {
            "$top": "100",  # Nombre d'emails par page
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,sender,bodyPreview,receivedDateTime,isRead,hasAttachments",
            "$expand": "to,cc"  # Pour récupérer les destinataires
        }

        # Ajout du filtre pour la synchronisation incrémentale
        if incremental:
            last_sync = self.get_last_sync_time()
            if last_sync:
                # Format ISO 8601 pour Microsoft Graph
                params["$filter"] = f"receivedDateTime gt {last_sync}"

        while True:
            try:
                # Construction de l'URL complète
                if next_link:
                    url = next_link
                else:
                    # Ajout des paramètres à l'URL
                    param_str = "&".join([f"{k}={v}" for k, v in params.items()])
                    url = f"{self.graph_base}/me/mailFolders/inbox/messages?{param_str}"

                print(f"🔍 Récupération de la page {page}...")
                response = requests.get(url, headers=self.headers)
                response.raise_for_status()
                data = response.json()

                # Traitement des emails
                emails = data.get("value", [])
                if not emails:
                    print("✅ Aucune nouvelle page d'emails")
                    break

                for email in emails:
                    # Vérification si l'email existe déjà
                    cursor.execute("""
                        SELECT id FROM emails
                        WHERE id = ? AND user_email = ?
                    """, (email["id"], self.user_email))
                    existing = cursor.fetchone()

                    # Préparation des données
                    recipients = {
                        "to": [{"email": t["emailAddress"]["address"], "name": t["emailAddress"]["name"]}
                              for t in email.get("toRecipients", [])],
                        "cc": [{"email": t["emailAddress"]["address"], "name": t["emailAddress"]["name"]}
                              for t in email.get("ccRecipients", [])]
                    }

                    # Insertion ou mise à jour
                    if existing:
                        # Mise à jour de l'email existant
                        cursor.execute("""
                            UPDATE emails SET
                                subject = ?,
                                sender = ?,
                                sender_email = ?,
                                recipients = ?,
                                date = ?,
                                body_preview = ?,
                                is_read = ?,
                                has_attachments = ?,
                                last_synced = CURRENT_TIMESTAMP
                            WHERE id = ? AND user_email = ?
                        """, (
                            email["subject"],
                            email["sender"]["emailAddress"]["name"],
                            email["sender"]["emailAddress"]["address"],
                            json.dumps(recipients),
                            email["receivedDateTime"],
                            email["bodyPreview"],
                            1 if email.get("isRead", False) else 0,
                            1 if email.get("hasAttachments", False) else 0,
                            email["id"],
                            self.user_email
                        ))
                        total_updated += 1
                    else:
                        # Insertion du nouvel email
                        cursor.execute("""
                            INSERT INTO emails (
                                id, subject, sender, sender_email, recipients,
                                date, body_preview, is_read, has_attachments, user_email
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            email["id"],
                            email["subject"],
                            email["sender"]["emailAddress"]["name"],
                            email["sender"]["emailAddress"]["address"],
                            json.dumps(recipients),
                            email["receivedDateTime"],
                            email["bodyPreview"],
                            1 if email.get("isRead", False) else 0,
                            1 if email.get("hasAttachments", False) else 0,
                            self.user_email
                        ))
                        total_fetched += 1

                # Gestion de la pagination
                next_link = data.get("@odata.nextLink")
                if not next_link or (max_emails and total_fetched + total_updated >= max_emails):
                    print(f"✅ Dernière page atteinte (total: {total_fetched + total_updated})")
                    break

                page += 1
                print(f"✅ Page {page} traitée. Total: {total_fetched + total_updated}")

            except Exception as e:
                print(f"❌ Erreur lors de la récupération des emails: {e}")
                import traceback
                traceback.print_exc()
                break

        conn.commit()
        conn.close()

        # Log de la synchronisation
        self._log_sync(total_fetched, total_updated)

        print(f"🎉 {total_fetched} nouveaux emails, {total_updated} mis à jour")
        return total_fetched + total_updated

    def search_emails(self, search_term, limit=10):
        """Recherche d'emails par sujet ou contenu"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute("""
            SELECT id, subject, sender, sender_email, date, body_preview
            FROM emails
            WHERE user_email = ?
            AND (subject LIKE ? OR body_preview LIKE ?)
            ORDER BY date DESC
            LIMIT ?
        """, (self.user_email, f"%{search_term}%", f"%{search_term}%", limit))

        results = cursor.fetchall()
        conn.close()

        return [{
            "id": row[0],
            "subject": row[1],
            "sender": row[2],
            "sender_email": row[3],
            "date": row[4],
            "body_preview": row[5]
        } for row in results]

    def get_email_body(self, email_id):
        """Récupère le corps complet d'un email"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute("""
            SELECT subject, sender, sender_email, date, body, recipients
            FROM emails
            WHERE id = ? AND user_email = ?
        """, (email_id, self.user_email))

        result = cursor.fetchone()
        conn.close()

        if result:
            return {
                "subject": result[0],
                "sender": result[1],
                "sender_email": result[2],
                "date": result[3],
                "body": result[4],
                "recipients": json.loads(result[5]) if result[5] else []
            }
        return None

    def get_stats(self):
        """Récupère les statistiques des emails"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Total d'emails
        cursor.execute("SELECT COUNT(*) FROM emails WHERE user_email = ?",
                       (self.user_email,))
        total_emails = cursor.fetchone()[0]

        # Emails non lus
        cursor.execute("SELECT COUNT(*) FROM emails WHERE user_email = ? AND is_read = 0",
                       (self.user_email,))
        unread_emails = cursor.fetchone()[0]

        # Emails avec pièces jointes
        cursor.execute("SELECT COUNT(*) FROM emails WHERE user_email = ? AND has_attachments = 1",
                       (self.user_email,))
        emails_with_attachments = cursor.fetchone()[0]

        # Dernière synchro
        cursor.execute("""
            SELECT sync_time, emails_fetched, emails_updated
            FROM sync_log
            ORDER BY sync_time DESC
            LIMIT 1
        """)
        sync_info = cursor.fetchone()

        conn.close()

        return {
            "total_emails": total_emails,
            "unread_emails": unread_emails,
            "emails_with_attachments": emails_with_attachments,
            "last_sync_time": sync_info[0] if sync_info else None,
            "last_sync_fetched": sync_info[1] if sync_info else None,
            "last_sync_updated": sync_info[2] if sync_info else None
        }
