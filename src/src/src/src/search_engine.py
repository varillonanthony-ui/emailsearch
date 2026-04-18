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
            return [{
                "id"          : r["id"],
                "subject"     : r["subject"],
                "sender"      : r["sender"],
                "sender_email": r["sender_email"],
                "date"        : r["date"],
                "body_preview": r["body_preview"],
                "score"       : r.score
            } for r in results]

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
