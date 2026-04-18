import sqlite3
from whoosh import index
from whoosh.qparser import MultifieldParser
from whoosh.query import And
import config

class SearchEngine:

    def __init__(self, db_path=None, index_path=None):
        self.db_path = db_path or config.DB_PATH
        self.index_path = index_path or config.INDEX_PATH
        try:
            self.ix = index.open_dir(self.index_path)
        except:
            self.ix = None
            print("⚠️ Index Whoosh non trouvé, utilisation de SQLite")

    def search(self, query_str, fields=None, limit=50):
        """
        Recherche multi-mots clés avec AND (tous les mots doivent être présents)
        Résultats triés par date (plus récent en haut)
        """
        if not fields:
            fields = ["subject", "body", "sender"]
        
        # ✅ Si Whoosh n'est pas disponible, utiliser SQLite
        if self.ix is None:
            return self._search_sqlite(query_str, limit)
        
        try:
            with self.ix.searcher() as searcher:
                parser = MultifieldParser(fields, self.ix.schema)
                
                # ✅ Split les mots et crée une requête AND (tous doivent être présents)
                keywords = query_str.split()
                if len(keywords) > 1:
                    # Créer une requête avec AND entre les mots
                    queries = []
                    for keyword in keywords:
                        try:
                            queries.append(parser.parse(keyword))
                        except:
                            pass
                    if queries:
                        query = And(queries)  # ✅ CHANGÉ DE Or À And
                    else:
                        query = parser.parse(query_str)
                else:
                    query = parser.parse(query_str)
                
                results = searcher.search(query, limit=limit * 2)
                ids = [r["id"] for r in results]
                scores = {r["id"]: r.score for r in results}
        except Exception as e:
            print(f"❌ Erreur Whoosh: {e}, utilisation de SQLite")
            return self._search_sqlite(query_str, limit)

        # Récupérer depuis SQLite
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        result_list = []
        for email_id in ids:
            cursor.execute(
                "SELECT id, subject, sender, sender_email, recipients, date, body_preview FROM emails WHERE id=?",
                (email_id,)
            )
            row = cursor.fetchone()
            if row:
                result_list.append({
                    "id": row[0],
                    "subject": row[1],
                    "sender": row[2],
                    "sender_email": row[3],
                    "to": row[4],
                    "date": row[5],
                    "body_preview": row[6],
                    "score": scores.get(row[0], 0)
                })
        conn.close()
        
        # ✅ TRIER PAR DATE (PLUS RÉCENT EN HAUT)
        result_list.sort(key=lambda x: x["date"] if x["date"] else "", reverse=True)
        
        # ✅ Limiter au nombre de résultats demandé
        return result_list[:limit]

    def _search_sqlite(self, query_str, limit=50):
        """
        Fallback : recherche directe dans SQLite avec LIKE
        ✅ TOUS les mots-clés doivent être présents (AND)
        """
        keywords = query_str.split()
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # ✅ CHANGÉ : Utiliser AND au lieu de OR
        # Chaque mot-clé doit être présent dans au moins UN champ
        where_clauses = []
        for keyword in keywords:
            where_clauses.append(f"(subject LIKE ? OR body LIKE ? OR sender LIKE ? OR sender_email LIKE ? OR recipients LIKE ?)")
        
        where_sql = " AND ".join(where_clauses)  # ✅ AND entre les mots
        
        params = []
        for keyword in keywords:
            for _ in range(5):  # 5 champs à vérifier
                params.append(f"%{keyword}%")
        
        # ✅ TRIER PAR DATE DESC (PLUS RÉCENT EN HAUT)
        cursor.execute(
            f"""SELECT id, subject, sender, sender_email, recipients, date, body_preview 
               FROM emails 
               WHERE {where_sql}
               ORDER BY date DESC
               LIMIT ?""",
            params + [limit]
        )
        
        result_list = []
        for row in cursor.fetchall():
            result_list.append({
                "id": row[0],
                "subject": row[1],
                "sender": row[2],
                "sender_email": row[3],
                "to": row[4],
                "date": row[5],
                "body_preview": row[6],
                "score": 0
            })
        conn.close()
        return result_list

    def get_email_detail(self, email_id):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM emails WHERE id=?", (email_id,))
        row = cursor.fetchone()
        conn.close()
        if row:
            return {
                "id": row[0],
                "subject": row[1],
                "sender": row[2],
                "sender_email": row[3],
                "recipients": row[4],
                "date": row[5],
                "body": row[6],
                "body_preview": row[7],
                "folder": row[8],
                "is_read": row[9],
                "has_attachments": row[10]
            }
        return None
