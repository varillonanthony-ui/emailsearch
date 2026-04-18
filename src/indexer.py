import sqlite3
import os
from whoosh import index
from whoosh.fields import Schema, TEXT, ID
from whoosh.analysis import StemmingAnalyzer
import config

class EmailIndexer:

    def __init__(self, db_path=None, index_path=None):
        """
        db_path: chemin vers la DB SQLite (optional)
        index_path: chemin vers l'index Whoosh (optional)
        """
        self.db_path = db_path or config.DB_PATH
        self.index_path = index_path or config.INDEX_PATH

        self.schema = Schema(
            id           = ID(stored=True, unique=True),
            subject      = TEXT(stored=True, analyzer=StemmingAnalyzer()),
            sender       = TEXT(stored=True),
            sender_email = ID(stored=True),
            date         = TEXT(stored=True),
            body         = TEXT(analyzer=StemmingAnalyzer()),
            body_preview = TEXT(stored=True),
            folder       = ID(stored=True),
        )

        os.makedirs(self.index_path, exist_ok=True)
        if index.exists_in(self.index_path):
            self.ix = index.open_dir(self.index_path)
        else:
            self.ix = index.create_in(self.index_path, self.schema)

    def index_all_emails(self):
        """Indexe tous les emails de la DB"""
        try:
            # ✅ Vérifie que la DB existe
            if not os.path.exists(self.db_path):
                raise FileNotFoundError(f"❌ DB non trouvée: {self.db_path}")

            conn   = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # ✅ Vérifie que la table existe
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='emails'")
            if not cursor.fetchone():
                raise Exception("❌ Table 'emails' n'existe pas dans la DB")
            
            # ✅ Récupère les emails
            cursor.execute(
                "SELECT id,subject,sender,sender_email,date,body,body_preview,folder FROM emails"
            )
            emails = cursor.fetchall()
            conn.close()

            if not emails:
                raise Exception("⚠️ Aucun email dans la DB! Récupère d'abord les emails.")

            print(f"📧 {len(emails)} emails trouvés, indexation en cours...")

            writer = self.ix.writer()
            for i, e in enumerate(emails):
                try:
                    writer.update_document(
                        id           = str(e[0]) if e[0] else "",
                        subject      = str(e[1]) if e[1] else "",
                        sender       = str(e[2]) if e[2] else "",
                        sender_email = str(e[3]) if e[3] else "",
                        date         = str(e[4]) if e[4] else "",
                        body         = str(e[5]) if e[5] else "",
                        body_preview = str(e[6]) if e[6] else "",
                        folder       = str(e[7]) if e[7] else "",
                    )
                except Exception as err:
                    print(f"⚠️ Erreur indexation email {e[0]}: {err}")
                    continue
                
                # Progress bar
                if (i + 1) % 500 == 0:
                    print(f"  ✅ {i + 1}/{len(emails)} indexés...")

            writer.commit()
            print(f"🎉 {len(emails)} emails indexés avec succès !")
            return len(emails)

        except Exception as err:
            print(f"❌ Erreur indexation: {err}")
            return 0
