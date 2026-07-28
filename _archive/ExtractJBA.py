import pyodbc
import sqlite3
import os
import time  # Added
from pathlib import Path
from config import SQLITE_DB
from config import DB2_CONN_STR
from config import BASE_DIR
import sqlite3
from tqdm import tqdm

QUERIES_DIR = BASE_DIR / "queries/jba"

def load_queries():
    queries = {}
    for filename in os.listdir(QUERIES_DIR):
        if filename.endswith(".sql"):
            name = filename[:-4]  # strip .sql
            with open(os.path.join(QUERIES_DIR, filename), 'r') as f:
                queries[name] = f.read()
    return queries

def fetch_db2_data(query):
    conn = pyodbc.connect(DB2_CONN_STR)
    cursor = conn.cursor()
    cursor.execute(query)
    columns = [col[0] for col in cursor.description]
    rows = cursor.fetchall()
    conn.close()
    print(f"Loaded {len(rows)} records.")
    return columns, rows

def recreate_sqlite_table(table_name, columns, rows):
    conn = sqlite3.connect(SQLITE_DB)
    cur = conn.cursor()

    cur.execute(f"DROP TABLE IF EXISTS {table_name}")
    col_defs = ", ".join([f"{col} TEXT" for col in columns])
    cur.execute(f"CREATE TABLE {table_name} ({col_defs})")
    placeholders = ", ".join(["?"] * len(columns))

    # Insert rows with progress bar
    print(f"Inserting {len(rows)} rows into {table_name}...")
    for row in tqdm(rows, desc=f"Inserting {table_name}", unit="row"):
        cur.execute(f"INSERT INTO {table_name} VALUES ({placeholders})", [str(item) for item in row])

    conn.commit()
    conn.close()

if __name__ == "__main__":
    start_time = time.time()  # Start timer
    queries = load_queries()
    for name, sql in queries.items():
        print(f"Running query: {name}.sql")
        columns, rows = fetch_db2_data(sql)
        recreate_sqlite_table(name.upper(), columns, rows)
        print(f"Loaded table {name.upper()} into SQLite.")
    elapsed = time.time() - start_time  # End timer
    hours, rem = divmod(elapsed, 3600)
    minutes, seconds = divmod(rem, 60)
    print(f"Total run time: {int(hours):02d}:{int(minutes):02d}:{seconds:02d} (hh:mm:ss)")