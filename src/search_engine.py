import sqlite3
from whoosh import index
from whoosh.qparser import MultifieldParser
import config

class SearchEngine:

    def __init__(self):
        self.ix = index.open_dir(config.INDEX_PATH)

    def search(self, query_str, fields=None, limit=50):
        if not fields:
            fields = ["subject", "body", "sender"]
        with self.ix.searcher() as searcher:
            parser  = MultifieldParser(fields, self.ix.schema)
            query   = parser.parse(query_str)
            results = searcher.search(query, limit=limit)
            ids     = [r["id"] for r in results]
            scores  = {r["id"]: r.score for r in results}

        # Récupérer les destinataires depuis SQLite
        conn = sqlite3.connect(config.DB_PATH)
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
                    "id"          : row[0],
                    "subject"     : row[1],
                    "sender"      : row[2],
                    "sender_email": row[3],
                    "to"          : row[4],
                    "date"        : row[5],
                    "body_preview": row[6],
                    "score"       : scores.get(row[0], 0)
                })
        conn.close()
        return result_list

    def get_email_detail(self, email_id):
        conn   = sqlite3.connect(config.DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM emails WHERE id=?", (email_id,))
        row = cursor.fetchone()
        conn.close()
        if row:
            return {
                "id"             : row[0],
                "subject"        : row[1],
                "sender"         : row[2],
                "sender_email"   : row[3],
                "recipients"     : row[4],
                "date"           : row[5],
                "body"           : row[6],
                "body_preview"   : row[7],
                "folder"         : row[8],
                "is_read"        : row[9],
                "has_attachments": row[10],
            }
        return None
