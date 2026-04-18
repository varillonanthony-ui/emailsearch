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

    def _get_indexed_ids(self):
        """Récupère les IDs des emails déjà indexés"""
        indexed_ids = set()
        try:
            searcher = self.ix.searcher()
            for doc in searcher.documents():
                indexed_ids.add(doc['id'])
            searcher.close()
        except Exception as e:
            print(f"⚠️ Erreur lecture index: {e}")
        return indexed_ids

    def index_all_emails(self, incremental=True):
        """
        Indexe tous les emails de la DB
        
        incremental=True: seulement les nouveaux emails
        incremental=False: ré-indexe TOUS les emails
        """
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
            all_emails = cursor.fetchall()
            conn.close()

            if not all_emails:
                raise Exception("⚠️ Aucun email dans la DB! Récupère d'abord les emails.")

            # ✅ MODE INCRÉMENTAL : filtre les emails déjà indexés
            if incremental:
                indexed_ids = self._get_indexed_ids()
                emails_to_index = [e for e in all_emails if str(e[0]) not in indexed_ids]
                
                if not emails_to_index:
                    print(f"✅ Aucun nouvel email à indexer ({len(all_emails)} emails en base)")
                    return 0
                
                print(f"📧 {len(emails_to_index)} NOUVEAUX emails à indexer ({len(all_emails)} au total)...")
            else:
                # ✅ MODE COMPLET : ré-indexe tout
                emails_to_index = all_emails
                print(f"📧 Ré-indexation de {len(emails_to_index)} emails...")

            # ✅ Indexe les emails
            writer = self.ix.writer()
            errors = 0
            
            for i, e in enumerate(emails_to_index):
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
                    errors += 1
                    continue

                # Progress bar tous les 500
                if (i + 1) % 500 == 0:
                    print(f"  ✅ {i + 1}/{len(emails_to_index)} indexés...")

            writer.commit()
            
            indexed_count = len(emails_to_index) - errors
            print(f"🎉 {indexed_count} emails indexés avec succès !")
            
            if errors > 0:
                print(f"⚠️ {errors} erreurs d'indexation")
            
            return indexed_count

        except Exception as err:
            print(f"❌ Erreur indexation: {err}")
            return 0

    def search(self, query, limit=50):
        """
        Recherche dans l'index Whoosh
        
        query: texte à chercher
        limit: nombre max de résultats
        """
        try:
            from whoosh.qparser import MultifieldParser
            
            searcher = self.ix.searcher()
            
            # ✅ Cherche dans plusieurs champs
            parser = MultifieldParser(
                ["subject", "body", "sender", "body_preview"],
                schema=self.schema
            )
            parsed_query = parser.parse(query)
            
            results = searcher.search(parsed_query, limit=limit)
            
            # ✅ Convertit les résultats
            hits = []
            for hit in results:
                hits.append({
                    "id": hit["id"],
                    "subject": hit["subject"],
                    "sender": hit["sender"],
                    "sender_email": hit["sender_email"],
                    "date": hit["date"],
                    "body_preview": hit["body_preview"],
                    "folder": hit["folder"],
                })
            
            searcher.close()
            return hits
            
        except Exception as e:
            print(f"❌ Erreur recherche: {e}")
            return []

    def get_index_stats(self):
        """Retourne les stats de l'index"""
        try:
            searcher = self.ix.searcher()
            count = searcher.doc_count_all()
            searcher.close()
            return {
                "total_indexed": count,
                "index_path": self.index_path
            }
        except Exception as e:
            print(f"⚠️ Erreur stats: {e}")
            return {"total_indexed": 0}
