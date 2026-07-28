import sqlite3

class APIBatchLogger:
    def __init__(self, db_path):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS api_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                    program TEXT,
                    transaction_name TEXT,
                    request_payload TEXT,
                    response_body TEXT,
                    status TEXT
                )
            """)

    def log(self, program, transaction_name, request, response, status="SUCCESS"):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                "INSERT INTO api_logs (program, transaction_name, request_payload, response_body, status) VALUES (?, ?, ?, ?, ?)",
                (program, transaction_name, json.dumps(request), json.dumps(response), status)
            )