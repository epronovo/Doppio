import sqlite3
import json

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
                    status TEXT,
                    nr_of_successful_transactions INTEGER,
                    nr_of_failed_transactions INTEGER
                )
            """)
            for col in ("nr_of_successful_transactions", "nr_of_failed_transactions"):
                try:
                    conn.execute(f"ALTER TABLE api_logs ADD COLUMN {col} INTEGER")
                except sqlite3.OperationalError:
                    pass  # column already exists

    def log(self, program, transaction_name, request, response, status="SUCCESS"):
        nr_success = response.get("nrOfSuccessfullTransactions") if isinstance(response, dict) else None
        nr_failed  = response.get("nrOfFailedTransactions")      if isinstance(response, dict) else None
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                "INSERT INTO api_logs (program, transaction_name, request_payload, response_body, status, nr_of_successful_transactions, nr_of_failed_transactions) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (program, transaction_name, json.dumps(request), json.dumps(response), status, nr_success, nr_failed)
            )