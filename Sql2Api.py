# Sql2Api.py
# -----------------------------------------------------------------------
# PURPOSE
#   Execute a SQL query against a local SQLite database and use the
#   resulting rows to drive batched M3 REST API calls.
#
#   COLUMN CONVENTION
#     col[0]  →  minm  : M3 program name  (e.g. "CRS610MI")
#     col[1]  →  trnm  : transaction name (e.g. "ChgFinancial")
#     col[2+] →  field key/value pairs that form the API record
#
#   The first two columns are consumed for routing; only col[2+] are
#   sent to M3.  Rows are grouped into batches (default 100, override
#   with --batch-size) and posted as a single JSON payload per batch.
#
# USAGE
#   python3 Sql2Api.py --sql-file  my_query.sql  [--batch-size 50]
#   python3 Sql2Api.py --sql       "SELECT ..."  [--batch-size 50]
#   python3 Sql2Api.py                           [--batch-size 50]
#         (no --sql / --sql-file → prompts for SQL interactively)
#
# DEPENDENCIES
#   InforMI.py, APIBatchLogger.py, UserDefaults.py, config.py
# -----------------------------------------------------------------------

import argparse
import json
import sqlite3
import sys
import threading
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from pathlib import Path

import requests
from tqdm import tqdm

from APIBatchLogger import APIBatchLogger
from InforMI import CONFIG, select_ionapi_file, get_ion_token
from UserDefaults import load_user_defaults, save_user_defaults

# -----------------------------------------------------------------------
# Defaults
# -----------------------------------------------------------------------
DEFAULT_BATCH_SIZE = 100
DEFAULT_WORKERS = 5
IONAPI_DIR = Path(__file__).parent / "ionapi"
SQLITE_DIR = Path.home() / "sqlite"

_token_lock = threading.Lock()


# -----------------------------------------------------------------------
# Startup / configuration helpers
# -----------------------------------------------------------------------

def prompt_for_tenant():
    """Select .ionapi file and acquire an OAuth token."""
    CONFIG["tenant"] = select_ionapi_file(IONAPI_DIR)
    get_ion_token()


def prompt_for_run_settings(batch_size_arg: int | None, workers_arg: int | None) -> dict:
    """
    Prompt for company, division, database name, batch size, and workers.
    Respects a 1-hour re-use window stored in user_defaults.json,
    mirroring the pattern in InforMI.prompt_for_company_division().
    Returns a dict with the resolved settings.
    """
    defaults = load_user_defaults()
    last_prompt = defaults.get("last_sql2api_prompt_time")
    now = datetime.now()

    # Re-use saved settings if prompted recently and no CLI overrides
    if last_prompt and batch_size_arg is None and workers_arg is None:
        last_dt = datetime.fromisoformat(last_prompt)
        if now - last_dt < timedelta(hours=1):
            settings = {
                "company":        defaults.get("company",        "100"),
                "division":       defaults.get("division",       ""),
                "local_db_name":  defaults.get("local_db_name",  "doppio.db"),
                "batch_size":     int(defaults.get("batch_size", DEFAULT_BATCH_SIZE)),
                "workers":        int(defaults.get("workers",    DEFAULT_WORKERS)),
            }
            _apply_settings(settings)
            return settings

    # Interactive prompts
    d_company   = defaults.get("company",       "100")
    d_division  = defaults.get("division",      "")
    d_db        = defaults.get("local_db_name", "doppio.db")
    d_batch     = defaults.get("batch_size",    DEFAULT_BATCH_SIZE)
    d_workers   = defaults.get("workers",       DEFAULT_WORKERS)

    db_name   = input(f"  SQLite database name  (default: {d_db}):       ").strip() or d_db
    company   = input(f"  Company code          (default: {d_company}):  ").strip() or d_company
    division  = input(f"  Division code         (default: {d_division or 'none'}): ").strip() or d_division
    if batch_size_arg is not None:
        batch_size = batch_size_arg
    else:
        bs_input = input(f"  Batch size            (default: {d_batch}):     ").strip()
        batch_size = int(bs_input) if bs_input else int(d_batch)
    if workers_arg is not None:
        workers = workers_arg
    else:
        w_input = input(f"  Workers               (default: {d_workers}):       ").strip()
        workers = int(w_input) if w_input else int(d_workers)

    settings = {
        "company":       company,
        "division":      division,
        "local_db_name": db_name,
        "batch_size":    batch_size,
        "workers":       workers,
    }

    # Persist
    defaults.update({
        "company":                  company,
        "division":                 division,
        "local_db_name":            db_name,
        "batch_size":               batch_size,
        "workers":                  workers,
        "last_sql2api_prompt_time": now.isoformat(),
    })
    save_user_defaults(defaults)
    _apply_settings(settings)
    print()
    return settings


def _apply_settings(settings: dict):
    """Push resolved settings into InforMI CONFIG and rebuild the API URL."""
    CONFIG["company"]  = settings["company"]
    CONFIG["division"] = settings["division"]

    divi_param = f"&divi={settings['division']}" if settings.get("division") else ""
    # CONFIG['iu'] and CONFIG['ti'] are set by get_ion_token() → we rebuild
    # api_url here so it picks up the correct cono/divi even when we skipped
    # the interactive prompt.
    if CONFIG.get("iu") and CONFIG.get("ti"):
        CONFIG["api_url"] = (
            f"{CONFIG['iu']}/{CONFIG['ti']}/M3/m3api-rest/v2/execute"
            f"?maxrecs=0&extendedresult=true&righttrim=true"
            f"&cono={settings['company']}{divi_param}"
        )


def get_sql_db_path(db_name: str) -> Path:
    path = SQLITE_DIR / db_name
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


# -----------------------------------------------------------------------
# SQL execution
# -----------------------------------------------------------------------

def run_sql(db_path: Path, sql: str) -> tuple[list[str], list[tuple]]:
    """
    Execute *sql* against the SQLite database at *db_path*.
    Returns (column_names, rows).
    """
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.execute(sql)
        rows = cur.fetchall()
        col_names = [d[0] for d in cur.description] if cur.description else []
    return col_names, rows


# -----------------------------------------------------------------------
# Payload construction
# -----------------------------------------------------------------------

def build_batches(col_names: list[str], rows, batch_size: int):
    """
    Group rows by (program, transaction) and yield API payloads in
    batches of *batch_size*.

    Each row must have at least 2 columns:
        col[0] → program  (minm)
        col[1] → transaction (trnm)
        col[2+] → record fields

    Yields dicts:  {"program": ..., "transactions": [...]}
    """
    if len(col_names) < 2:
        raise ValueError(
            "SQL must return at least 2 columns: [minm] and [trnm]. "
            f"Got: {col_names}"
        )

    field_names = col_names[2:]  # field columns start at index 2

    # Group records by (program, transaction)
    grouped: dict[tuple, list[dict]] = defaultdict(list)
    for row in rows:
        program     = str(row[0]).strip()
        transaction = str(row[1]).strip()
        record      = {
            field_names[i]: str(row[i + 2]).strip()
            for i in range(len(field_names))
            if row[i + 2] is not None and str(row[i + 2]).strip() != ""
        }
        grouped[(program, transaction)].append(record)

    # Yield payloads in batch_size chunks
    for (program, transaction), records in grouped.items():
        for offset in range(0, len(records), batch_size):
            batch = records[offset: offset + batch_size]
            yield {
                "program": program,
                "transactions": [
                    {"transaction": transaction, "record": rec}
                    for rec in batch
                ]
            }


# -----------------------------------------------------------------------
# M3 API posting (mirrors InforMI.post_to_m3 with inline session)
# -----------------------------------------------------------------------

def post_payload(payload: dict, session: requests.Session,
                 api_logger: APIBatchLogger,
                 log_lock: threading.Lock,
                 max_retries: int = 3, retry_delay: int = 2) -> dict:
    """
    POST a single payload to the M3 API with retry / token-refresh logic.
    Logs every attempt via APIBatchLogger. Thread-safe.
    """
    program  = payload.get("program", "UNKNOWN")
    tx_name  = payload.get("transactions", [{}])[0].get("transaction", "UNKNOWN")

    for attempt in range(1, max_retries + 1):
        try:
            headers = {
                "Authorization": f"Bearer {CONFIG['access_token']}",
                "Content-Type":  "application/json",
                "Accept":        "application/json; charset=UTF-8",
            }
            response = session.post(
                CONFIG["api_url"], json=payload, headers=headers, timeout=300
            )

            if response.status_code == 401:
                tqdm.write("  🔄  Token expired — refreshing …")
                with _token_lock:
                    get_ion_token()
                continue

            response.raise_for_status()
            res_json = response.json()
            with log_lock:
                api_logger.log(program, tx_name, payload, res_json, "SUCCESS")
            return res_json

        except Exception as exc:
            with log_lock:
                api_logger.log(program, tx_name, payload, {"error": str(exc)},
                               f"FAILED_ATTEMPT_{attempt}")
            if attempt == max_retries:
                raise RuntimeError(f"M3 API call failed after {max_retries} attempts: {exc}")
            tqdm.write(f"  ⚠  Attempt {attempt} failed: {exc} — retrying in {retry_delay}s …")
            time.sleep(retry_delay)


# -----------------------------------------------------------------------
# Result summary helpers
# -----------------------------------------------------------------------

def _count_errors(res_json: dict) -> int:
    """Count error entries returned inside an extendedresult payload."""
    try:
        results = res_json.get("results", [])
        return sum(
            1 for r in results
            if r.get("errorMessage") or r.get("errorCode")
        )
    except Exception:
        return 0


# -----------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Execute a SQL query and pipe results into M3 REST API calls.\n"
            "Column 1 = program (minm), Column 2 = transaction (trnm), "
            "remaining columns = field values."
        )
    )
    parser.add_argument(
        "--sql", "-s",
        metavar="SQL",
        help="Inline SQL string to execute"
    )
    parser.add_argument(
        "--sql-file", "-f",
        metavar="FILE",
        help="Path to a .sql file containing the query"
    )
    parser.add_argument(
        "--batch-size", "-b",
        metavar="N",
        type=int,
        default=None,
        help=f"Records per API call (default: {DEFAULT_BATCH_SIZE})"
    )
    parser.add_argument(
        "--workers", "-w",
        metavar="N",
        type=int,
        default=None,
        help=f"Parallel API workers (default: {DEFAULT_WORKERS})"
    )
    parser.add_argument(
        "--dry-run", "-n",
        action="store_true",
        help="Build and print payloads without calling the API"
    )
    args = parser.parse_args()

    print("\n🔷  Sql2Api — SQL → M3 REST API batch processor")
    print("=" * 52)

    # ------------------------------------------------------------------
    # 1. Resolve SQL
    # ------------------------------------------------------------------
    sql: str | None = None

    if args.sql:
        sql = args.sql.strip()
    elif args.sql_file:
        path = Path(args.sql_file)
        if not path.is_file():
            print(f"❌  SQL file not found: {path}")
            sys.exit(1)
        sql = path.read_text(encoding="utf-8").strip()
    else:
        print("\nNo SQL provided via --sql or --sql-file.")
        print("Enter your SQL query below (type END on a new line to finish):\n")
        lines = []
        while True:
            line = input()
            if line.strip().upper() == "END":
                break
            lines.append(line)
        sql = "\n".join(lines).strip()

    if not sql:
        print("❌  No SQL provided. Exiting.")
        sys.exit(1)

    # ------------------------------------------------------------------
    # 2. Auth + settings
    # ------------------------------------------------------------------
    print("\n🔐  Authentication")
    prompt_for_tenant()
    print("  ✅  Token acquired\n")

    print("⚙️   Run settings")
    settings = prompt_for_run_settings(args.batch_size, args.workers)
    batch_size = settings["batch_size"]
    workers = settings["workers"]

    # Rebuild api_url now that we have full CONFIG
    _apply_settings(settings)

    # ------------------------------------------------------------------
    # 3. Run SQL
    # ------------------------------------------------------------------
    db_path = get_sql_db_path(settings["local_db_name"])
    print(f"🗄️   Database : {db_path}")
    print(f"📝  SQL      :\n{sql}\n")

    col_names, rows = run_sql(db_path, sql)
    print(f"✅  Rows returned : {len(rows)}")

    if not rows:
        print("⚠️   No rows — nothing to send.")
        sys.exit(0)

    # ------------------------------------------------------------------
    # 4. Build payloads
    # ------------------------------------------------------------------
    payloads = list(build_batches(col_names, rows, batch_size))
    total_batches = len(payloads)
    total_records = sum(
        len(p["transactions"]) for p in payloads
    )
    print(f"📦  Batches : {total_batches}  ({total_records} records, batch size {batch_size})")

    # ------------------------------------------------------------------
    # 5. Dry run — just print and exit
    # ------------------------------------------------------------------
    if args.dry_run:
        print("\n🟡  DRY RUN — payloads (not sent):\n")
        for i, p in enumerate(payloads, 1):
            print(f"--- Batch {i} ---")
            print(json.dumps(p, indent=2))
        sys.exit(0)

    # ------------------------------------------------------------------
    # 6. POST to M3
    # ------------------------------------------------------------------
    api_logger = APIBatchLogger(str(get_sql_db_path(settings["local_db_name"])))
    log_lock = threading.Lock()

    success_count = 0
    error_count   = 0
    total_api_errors = 0

    print(f"\n🚀  Posting to M3  [{CONFIG['api_url']}]  (workers: {workers})\n")

    with requests.Session() as session:
        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {
                executor.submit(post_payload, payload, session, api_logger, log_lock): payload
                for payload in payloads
            }
            with tqdm(total=total_batches, desc="Batches", unit="batch") as pbar:
                for future in as_completed(futures):
                    payload = futures[future]
                    program   = payload["program"]
                    tx_name   = payload["transactions"][0]["transaction"]
                    rec_count = len(payload["transactions"])
                    try:
                        result = future.result()
                        api_errors = _count_errors(result)
                        total_api_errors += api_errors
                        success_count += 1
                        if api_errors:
                            tqdm.write(
                                f"  ⚠  {program}/{tx_name}  {rec_count} recs  "
                                f"— {api_errors} API-level error(s)"
                            )
                    except RuntimeError as exc:
                        error_count += 1
                        tqdm.write(f"  ❌  {program}/{tx_name}  {rec_count} recs  — {exc}")
                    pbar.update(1)

    # ------------------------------------------------------------------
    # 7. Summary
    # ------------------------------------------------------------------
    print("\n" + "=" * 52)
    print(f"✅  Successful batches : {success_count}/{total_batches}")
    if error_count:
        print(f"❌  Failed batches     : {error_count}/{total_batches}")
    if total_api_errors:
        print(f"⚠️   M3 record errors   : {total_api_errors}")
    print(f"📋  Log saved to       : {get_sql_db_path(settings['local_db_name'])}")
    print("=" * 52 + "\n")


if __name__ == "__main__":
    main()
