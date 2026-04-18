import sqlite3
import os
from whoosh import index
from whoosh.fields import Schema, TEXT, ID
from whoosh.analysis import StemmingAnalyzer
import config

class EmailIndexer:

    def __init__(self):
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
        os.makedirs(config.INDEX_PATH, exist_ok=True)
        if index.exists_in(config.INDEX_PATH):
            self.ix = index.open_dir(config.INDEX_PATH)
        else:
            self.ix = index.create_in(config.INDEX_PATH, self.schema)

    def index_all_emails(self):
        conn   = sqlite3.connect(config.DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id,subject,sender,sender_email,date,body,body_preview,folder FROM emails"
        )
        emails = cursor.fetchall()
        conn.close()

        writer = self.ix.writer()
        for e in emails:
            writer.update_document(
                id           = e[0],
                subject      = e[1] or "",
                sender       = e[2] or "",
                sender_email = e[3] or "",
                date         = e[4] or "",
                body         = e[5] or "",
                body_preview = e[6] or "",
                folder       = e[7] or "",
            )
        writer.commit()
        return len(emails)
