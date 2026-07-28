# ExtractJBA.py 
import os
import re
import time
import pyodbc
import sqlite3

from decimal import Decimal
from pathlib import Path
from tqdm import tqdm
from InforMI import prompt_for_company_division
from config import DB2_CONN_STR, BASE_DIR, get_sqlite_db_path

# initialize field
print(f"")
QUERIES_DIR = BASE_DIR / "queries/jba"
BATCH_SIZE = 500  # Lower if needed
prompt_for_company_division()
SQLITE_DB_PATH = get_sqlite_db_path()

def load_queries():
    queries = {}
    for filename in os.listdir(QUERIES_DIR):
        if filename.endswith(".sql"):
            name = re.sub(r'_\[.*\]', '', filename[:-4])
            full_path = os.path.join(QUERIES_DIR, filename)
            with open(full_path, "r") as f:
                queries[name] = (f.read(), full_path)
    return queries

def extract_sqlite_subquery(sql):
    pattern = r"IN\s*\(\s*(SELECT\s+.*?\s+FROM\s+sqlite\.[^)]+?)\s*\)"
    match = re.search(pattern, sql, flags=re.IGNORECASE | re.DOTALL)
    if not match:
        return sql, None
    subquery = match.group(1)
    modified_sql = sql[:match.start()] + "IN ({})" + sql[match.end():]
    return modified_sql, subquery


def run_sqlite_subquery(subquery, truncate_len=15):
    # print(f"Running SQLite subquery: {subquery}")
    sqlite_conn = sqlite3.connect(SQLITE_DB_PATH)
    sqlite_cursor = sqlite_conn.cursor()
    sqlite_cursor.execute(subquery.replace("sqlite.", ""))
    rows = [row[0][:truncate_len] for row in sqlite_cursor.fetchall() if row[0]]
    sqlite_conn.close()
    # print(f"Loaded {len(rows)} values from SQLite.")
    return rows


def fetch_db2_data(query, batch_params=None):
    conn = pyodbc.connect(DB2_CONN_STR)
    cursor = conn.cursor()
    all_rows = []
    column_names = None

    if batch_params:
        # print(f"Querying DB2 in batches of {BATCH_SIZE}...")
        for i in tqdm(range(0, len(batch_params), BATCH_SIZE), desc="DB2 Batches", unit="batch"):
            batch = batch_params[i:i + BATCH_SIZE]
            placeholders = ",".join("?" for _ in batch)
            batched_query = query.format(placeholders)
            cursor.execute(batched_query, batch)
            if column_names is None:
                column_names = [desc[0] for desc in cursor.description]
            all_rows.extend(cursor.fetchall())
    else:
        # print(f"Executing query with streaming fetch (batch size = {BATCH_SIZE})...")
        cursor.execute(query)
        column_names = [desc[0] for desc in cursor.description]

        with tqdm(desc="Fetching rows", unit="rows") as pbar:
            while True:
                rows = cursor.fetchmany(BATCH_SIZE)
                if not rows:
                    break
                all_rows.extend(rows)
                pbar.update(len(rows))

    conn.close()
    return column_names, all_rows


def recreate_sqlite_table(table_name, column_names, rows):
    sqlite_conn = sqlite3.connect(SQLITE_DB_PATH)
    sqlite_cursor = sqlite_conn.cursor()

    # print(f"Dropping existing table {table_name} if it exists...")
    sqlite_cursor.execute(f"DROP TABLE IF EXISTS {table_name}")

    quoted_column_names = [f'"{col}"' for col in column_names]
    columns_sql = ", ".join([f'{col} TEXT' for col in quoted_column_names])
    sqlite_cursor.execute(f"CREATE TABLE {table_name} ({columns_sql})")

    insert_sql = f'INSERT INTO {table_name} ({", ".join(quoted_column_names)}) VALUES ({", ".join(["?"] * len(column_names))})'

    cleaned_rows = []
    for row in tqdm(rows, desc=f"Inserting into {table_name}", unit="row"):
        cleaned_row = []
        for val in row:
            if isinstance(val, Decimal):
                cleaned_row.append(str(val))  # or float(val)
            else:
                cleaned_row.append(val)
        cleaned_rows.append(tuple(cleaned_row))

    sqlite_cursor.executemany(insert_sql, cleaned_rows)
    sqlite_conn.commit()
    sqlite_conn.close()
    # print(f"✅ Finished writing {table_name} to SQLite.")


if __name__ == "__main__":
    start = time.time()
    queries = load_queries()

    archive_dir = QUERIES_DIR / "_archive"
    archive_dir.mkdir(exist_ok=True)

    for name, (sql, file_path) in queries.items():
        print(f"🔍 Processing: {name}.sql")
        modified_sql, sqlite_subquery = extract_sqlite_subquery(sql)

        if sqlite_subquery:
            values = run_sqlite_subquery(sqlite_subquery)
            if not values:
                print(f"⚠️ No values found for {name}, skipping.")
                continue
            column_names, rows = fetch_db2_data(modified_sql, batch_params=values)
        else:
            column_names, rows = fetch_db2_data(sql)

        recreate_sqlite_table(name.upper(), column_names, rows)

        # ✅ Move file to _archive after processing
        archived_path = archive_dir / Path(file_path).name
        os.rename(file_path, archived_path)
        print(f"📦 Archived: {archived_path}")

    elapsed = time.time() - start
    h, rem = divmod(elapsed, 3600)
    m, s = divmod(rem, 60)
    print(f"✅ Done. Total time: {int(h):02d}:{int(m):02d}:{int(s):02d}")