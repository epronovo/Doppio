"""
Teamwork_SetupSync.py
---------------------
Creates the BigQuery dataset and all tables used by Teamwork_Contracts.py and
Teamwork_TimeData.py, then copies all data from the local SQLite database.

Safe to re-run: tables are created with exists_ok, existing BQ data is replaced
with WRITE_TRUNCATE so SQLite is always the source of truth.

BigQuery: spatial-earth-492100-b7.teamwork
SQLite  : teamwork.db
Tables  : ProjectCustomFields, Contracts, ContractFieldValues, TimeData, SyncState
"""

import re
import sqlite3
import numpy as np
import pandas as pd
from datetime import datetime
from google.cloud import bigquery
from google.oauth2 import service_account

# ── Config ────────────────────────────────────────────────────────────────────
BQ_PROJECT = 'spatial-earth-492100-b7'
BQ_DATASET = 'teamwork'
KEY_PATH   = '/Users/ericpronovost/Downloads/spatial-earth-492100-b7-cd5d8fb255b4.json'
DB_FILE    = 'teamwork.db'

# TimeData: SQLite uses display names ("Project ID"), BQ uses underscored ("Project_ID")
_TIMEDATA_RENAME = {
    col: re.sub(r'[^A-Za-z0-9_]', '_', col)
    for col in [
        'Project ID', 'Project Name', 'Customer Name', 'Task List Name',
        'Task Name', 'Is subtask (Yes/No)', 'Parent Task Name', 'Date',
        'Billable Hours', 'Non Billable Hours', 'Billed Hours', 'Unbilled Hours',
        'Task Estimated Time', 'Person ID', 'First Name', 'Last Name', 'Task ID',
    ]
}

# ── Table schemas ─────────────────────────────────────────────────────────────
TABLES = {
    "ProjectCustomFields": [
        bigquery.SchemaField("project_id",  "INTEGER"),
        bigquery.SchemaField("field_id",    "STRING"),
        bigquery.SchemaField("field_name",  "STRING"),
        bigquery.SchemaField("field_type",  "STRING"),
        bigquery.SchemaField("value",       "STRING"),
    ],
    "Contracts": [
        bigquery.SchemaField("project_id",  "INTEGER"),
        bigquery.SchemaField("custom_item", "STRING"),
        bigquery.SchemaField("section",     "STRING"),
        bigquery.SchemaField("record_name", "STRING"),
        bigquery.SchemaField("fields_json", "STRING"),
        bigquery.SchemaField("fetched_at",  "TIMESTAMP"),
    ],
    "ContractFieldValues": [
        bigquery.SchemaField("project_id",  "INTEGER"),
        bigquery.SchemaField("section",     "STRING"),
        bigquery.SchemaField("record_name", "STRING"),
        bigquery.SchemaField("fetched_at",  "TIMESTAMP"),
        # Dynamic contract field columns are added at runtime by Teamwork_Contracts.py
    ],
    "TimeData": [
        bigquery.SchemaField("id",                    "INTEGER"),
        bigquery.SchemaField("Project_ID",            "INTEGER"),
        bigquery.SchemaField("Project_Name",          "STRING"),
        bigquery.SchemaField("Customer_Name",         "STRING"),
        bigquery.SchemaField("Task_List_Name",        "STRING"),
        bigquery.SchemaField("Task_Name",             "STRING"),
        bigquery.SchemaField("Is_subtask__Yes_No_",   "BOOLEAN"),
        bigquery.SchemaField("Parent_Task_Name",      "STRING"),
        bigquery.SchemaField("Date",                  "DATE"),
        bigquery.SchemaField("Billable_Hours",        "FLOAT"),
        bigquery.SchemaField("Non_Billable_Hours",    "FLOAT"),
        bigquery.SchemaField("Billed_Hours",          "FLOAT"),
        bigquery.SchemaField("Unbilled_Hours",        "FLOAT"),
        bigquery.SchemaField("Task_Estimated_Time",   "FLOAT"),
        bigquery.SchemaField("Person_ID",             "INTEGER"),
        bigquery.SchemaField("First_Name",            "STRING"),
        bigquery.SchemaField("Last_Name",             "STRING"),
        bigquery.SchemaField("Task_ID",               "INTEGER"),
    ],
    "SyncState": [
        bigquery.SchemaField("name",      "STRING"),
        bigquery.SchemaField("last_sync", "STRING"),
    ],
}


# ── BQ helpers ────────────────────────────────────────────────────────────────

def get_bq_client():
    creds = service_account.Credentials.from_service_account_file(
        KEY_PATH, scopes=["https://www.googleapis.com/auth/bigquery"])
    return bigquery.Client(project=BQ_PROJECT, credentials=creds)


def _setup_bq(client):
    dataset_ref = bigquery.Dataset(f"{BQ_PROJECT}.{BQ_DATASET}")
    dataset_ref.location = "US"
    client.create_dataset(dataset_ref, exists_ok=True)
    print(f"Dataset ready: {BQ_PROJECT}.{BQ_DATASET}\n")

    for table_name, schema in TABLES.items():
        table = bigquery.Table(f"{BQ_PROJECT}.{BQ_DATASET}.{table_name}", schema=schema)
        client.create_table(table, exists_ok=True)
        print(f"  Table ready: {table_name}")
    print()


def _ensure_bq_columns(client, table_name: str, extra_cols: set):
    table_ref = client.dataset(BQ_DATASET).table(table_name)
    table     = client.get_table(table_ref)
    existing  = {f.name for f in table.schema}
    new_fields = [
        bigquery.SchemaField(col, "STRING")
        for col in sorted(extra_cols)
        if col not in existing
    ]
    if new_fields:
        table.schema = table.schema + new_fields
        client.update_table(table, ["schema"])
        print(f"    Added {len(new_fields)} dynamic column(s): {[f.name for f in new_fields]}")


def _clean_rows(df: pd.DataFrame) -> list[dict]:
    result = []
    for rec in df.to_dict(orient='records'):
        clean = {}
        for k, v in rec.items():
            if isinstance(v, np.integer):
                clean[k] = int(v)
            elif isinstance(v, np.floating):
                clean[k] = None if np.isnan(v) else float(v)
            elif isinstance(v, np.bool_):
                clean[k] = bool(v)
            elif isinstance(v, float) and np.isnan(v):
                clean[k] = None
            else:
                clean[k] = v
        result.append(clean)
    return result


def _load_to_bq(client, table_name: str, rows: list[dict]):
    if not rows:
        print(f"    (no data)")
        return
    job_config = bigquery.LoadJobConfig(
        write_disposition=bigquery.WriteDisposition.WRITE_TRUNCATE,
        source_format=bigquery.SourceFormat.NEWLINE_DELIMITED_JSON,
    )
    client.load_table_from_json(
        rows, f"{BQ_PROJECT}.{BQ_DATASET}.{table_name}", job_config=job_config
    ).result()
    print(f"    Loaded {len(rows):,} rows → {BQ_PROJECT}.{BQ_DATASET}.{table_name}")


# ── Per-table copy functions ───────────────────────────────────────────────────

def _read_sqlite(conn, sqlite_table: str) -> pd.DataFrame | None:
    try:
        return pd.read_sql(f'SELECT * FROM "{sqlite_table}"', conn)
    except Exception as e:
        print(f"    Skipped — not found in SQLite: {e}")
        return None


def _copy_project_custom_fields(client, conn):
    print("  ProjectCustomFields:")
    df = _read_sqlite(conn, 'ProjectCustomFields')
    if df is not None:
        print(f"    {len(df):,} rows")
        _load_to_bq(client, 'ProjectCustomFields', _clean_rows(df))


def _copy_contracts(client, conn):
    print("  Contracts:")
    # SQLite table name is lowercase "contracts"
    df = _read_sqlite(conn, 'contracts')
    if df is not None:
        print(f"    {len(df):,} rows")
        _load_to_bq(client, 'Contracts', _clean_rows(df))


def _copy_contract_field_values(client, conn):
    print("  ContractFieldValues:")
    df = _read_sqlite(conn, 'ContractFieldValues')
    if df is None:
        return
    print(f"    {len(df):,} rows")
    base_cols = {'project_id', 'section', 'record_name', 'fetched_at'}
    extra_cols = {c for c in df.columns if c not in base_cols}
    if extra_cols:
        _ensure_bq_columns(client, 'ContractFieldValues', extra_cols)
    _load_to_bq(client, 'ContractFieldValues', _clean_rows(df))


def _copy_time_data(client, conn):
    print("  TimeData:")
    df = _read_sqlite(conn, 'TimeData')
    if df is None:
        return
    print(f"    {len(df):,} rows")
    df = df.rename(columns=_TIMEDATA_RENAME)
    bool_col = 'Is_subtask__Yes_No_'
    if bool_col in df.columns:
        df[bool_col] = df[bool_col].apply(
            lambda v: None if pd.isna(v) else bool(int(v)))
    _load_to_bq(client, 'TimeData', _clean_rows(df))


def _copy_sync_state(client, conn):
    print("  SyncState:")
    df = _read_sqlite(conn, 'SyncState')
    if df is not None:
        print(f"    {len(df):,} rows")
        _load_to_bq(client, 'SyncState', _clean_rows(df))


# ── Main ──────────────────────────────────────────────────────────────────────

def start_function():
    start_timestamp = datetime.now()
    print(f"Script Started: {start_timestamp}\n")

    client = get_bq_client()
    conn   = sqlite3.connect(DB_FILE)

    print("Setting up BigQuery...")
    _setup_bq(client)

    print("Copying data from SQLite → BigQuery...")
    _copy_project_custom_fields(client, conn)
    _copy_contracts(client, conn)
    _copy_contract_field_values(client, conn)
    _copy_time_data(client, conn)
    _copy_sync_state(client, conn)

    conn.close()

    end_timestamp = datetime.now()
    print(f"\nDone.")
    print(f"Script Finished: {end_timestamp}")
    print(f"Run Duration (in seconds): {(end_timestamp - start_timestamp).total_seconds()}")


if __name__ == "__main__":
    start_function()
