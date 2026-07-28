"""
Local_Contracts.py
------------------
SQLite-only version of Teamwork_Contracts.py for local development.
Writes to teamwork.db — no BigQuery dependency.
"""

import json
import re
import sqlite3
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone

import requests

# ── Teamwork ──────────────────────────────────────────────────────────────────
HOST    = 'https://doppiogroup.teamwork.com'
API_KEY = 'ZXJpY0Bkb3BwaW9ncm91cC5jb206WnNlNDVyZFhET1BQSU8wMQ=='
HEADERS = {"Authorization": "Basic " + API_KEY}

# ── SQLite ────────────────────────────────────────────────────────────────────
DB_FILE = 'teamwork.db'

_SEP         = '\x1f'
_API_WORKERS = 2


# ── Teamwork API helpers ──────────────────────────────────────────────────────

def get_json(url: str):
    for attempt in range(5):
        response = requests.get(f"{HOST}{url}", headers=HEADERS)
        if response.status_code == 429:
            wait = 2 ** attempt
            print(f"  Rate limited, retrying in {wait}s...")
            time.sleep(wait)
            continue
        response.raise_for_status()
        return response.json()
    response.raise_for_status()


def fetch_all_projects() -> list[dict]:
    projects = []
    page = 1
    while True:
        data = get_json(
            f'/projects/api/v3/projects.json'
            f'?projectStatuses=all&includeArchivedProjects=true'
            f'&includeCustomFields=true&pageSize=500&page={page}'
        )
        batch = data.get('projects', [])
        if not batch:
            break

        included  = data.get('included', {})
        raw_cfs   = included.get('customfields', [])
        cf_list   = raw_cfs.values() if isinstance(raw_cfs, dict) else raw_cfs
        cf_defs   = {cf['id']: cf for cf in cf_list}

        raw_cfp        = included.get('customfieldProjects', [])
        cf_values_list = raw_cfp.values() if isinstance(raw_cfp, dict) else raw_cfp

        cf_by_project: dict[int, list] = {}
        for cfv in cf_values_list:
            pid = cfv.get('projectId') or cfv.get('project', {}).get('id')
            if pid:
                cf_by_project.setdefault(pid, []).append(cfv)

        for project in batch:
            pid = project['id']
            project['customFields'] = [
                {
                    'id':    cfv.get('customfieldId'),
                    'name':  cf_defs.get(cfv.get('customfieldId'), {}).get('name'),
                    'type':  cf_defs.get(cfv.get('customfieldId'), {}).get('type'),
                    'value': cfv.get('value'),
                }
                for cfv in cf_by_project.get(pid, [])
            ]

        projects.extend(batch)

        meta = data.get('meta', {}).get('page', {})
        print(f"  Fetched {len(projects)} projects...")
        if not meta.get('hasMore', False):
            break
        page += 1

    return projects


def fetch_company_names(company_ids: list[int]) -> dict[int, str]:
    if not company_ids:
        return {}
    ids_str = ','.join(str(i) for i in company_ids)
    data = get_json(f'/projects/api/v3/companies.json?ids={ids_str}&pageSize=500')
    return {c['id']: c.get('name', '') for c in data.get('companies', [])}


def fetch_contracts(project_id: int) -> list[dict]:
    items_data   = get_json(f'/projects/api/v3/projects/{project_id}/customitems.json')
    custom_items = items_data.get('customItems', [])
    if not custom_items:
        return []

    with ThreadPoolExecutor(max_workers=min(12, len(custom_items) * 3)) as ex:
        f_fields   = [ex.submit(get_json, f'/projects/api/v3/customitems/{i["id"]}/fields.json')   for i in custom_items]
        f_sections = [ex.submit(get_json, f'/projects/api/v3/customitems/{i["id"]}/sections.json') for i in custom_items]
        f_records  = [ex.submit(get_json, f'/projects/api/v3/customitems/{i["id"]}/records.json')  for i in custom_items]
        all_fields   = [f.result() for f in f_fields]
        all_sections = [f.result() for f in f_sections]
        all_records  = [f.result() for f in f_records]

    contracts = []
    for item, fields_data, sections_data, records_data in zip(
        custom_items, all_fields, all_sections, all_records
    ):
        field_map = {
            f['twId']: f['displayName']
            for f in fields_data.get('customItemFields', [])
            if f.get('twId')
        }
        section_map = {
            s['id']: s['displayName']
            for s in sections_data.get('customItemSections', [])
        }
        for rec in records_data.get('customItemRecords', []):
            named_values = {
                field_map.get(uuid, uuid): value
                for uuid, value in (rec.get('fieldValues') or {}).items()
            }
            section_id   = (rec.get('section') or {}).get('id')
            section_name = section_map.get(section_id, '')
            contracts.append({
                'record':     rec['name'],
                'section':    section_name,
                'customItem': item.get('displayName', str(item['id'])),
                'fields':     named_values,
            })

    return contracts


# ── Column sanitizer ──────────────────────────────────────────────────────────

def _col(name: str) -> str:
    col = re.sub(r'[^A-Za-z0-9_]', '_', name)
    if col and col[0].isdigit():
        col = '_' + col
    return col or '_unknown'


# ── SQLite helpers ────────────────────────────────────────────────────────────

def init_tables(conn: sqlite3.Connection):
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS ProjectCustomFields (
            project_id  INTEGER,
            field_id    TEXT,
            field_name  TEXT,
            field_type  TEXT,
            value       TEXT,
            PRIMARY KEY (project_id, field_id)
        );
        CREATE TABLE IF NOT EXISTS Contracts (
            project_id   INTEGER,
            custom_item  TEXT,
            section      TEXT,
            record_name  TEXT,
            fields_json  TEXT,
            fetched_at   TEXT,
            company_name TEXT,
            PRIMARY KEY (project_id, custom_item, section, record_name)
        );
        CREATE TABLE IF NOT EXISTS ContractFieldValues (
            project_id   INTEGER,
            project_name TEXT,
            section      TEXT,
            record_name  TEXT,
            fetched_at   TEXT,
            company_name TEXT,
            PRIMARY KEY (project_id, section, record_name)
        );
        CREATE TABLE IF NOT EXISTS SyncState (
            name      TEXT PRIMARY KEY,
            last_sync TEXT
        );
    """)
    conn.commit()


def _ensure_columns(conn: sqlite3.Connection, table: str, col_names: set[str]):
    cursor   = conn.execute(f"PRAGMA table_info({table})")
    existing = {row[1] for row in cursor.fetchall()}
    new_cols = col_names - existing
    for col in sorted(new_cols):
        conn.execute(f'ALTER TABLE {table} ADD COLUMN "{col}" TEXT')
    if new_cols:
        conn.commit()
        print(f"  Added {len(new_cols)} column(s) to {table}")


def _load_existing_pcf(conn: sqlite3.Connection, project_ids: list[int]) -> dict:
    if not project_ids:
        return {}
    placeholders = ','.join('?' * len(project_ids))
    cursor = conn.execute(
        f"SELECT project_id, field_id, field_name, field_type, value "
        f"FROM ProjectCustomFields WHERE project_id IN ({placeholders})",
        project_ids,
    )
    result: dict[int, dict] = {}
    for pid, fid, fname, ftype, val in cursor.fetchall():
        result.setdefault(pid, {})[fid] = {
            'project_id': pid, 'field_id': fid,
            'field_name': fname or '', 'field_type': ftype or '', 'value': val or '',
        }
    return result


def _load_existing_contracts(conn: sqlite3.Connection, project_ids: list[int]) -> dict:
    if not project_ids:
        return {}
    placeholders = ','.join('?' * len(project_ids))
    cursor = conn.execute(
        f"SELECT project_id, custom_item, section, record_name, fields_json, company_name "
        f"FROM Contracts WHERE project_id IN ({placeholders})",
        project_ids,
    )
    result: dict[int, dict] = {}
    for pid, ci, sec, rn, fj, cn in cursor.fetchall():
        k = f"{ci}{_SEP}{sec}{_SEP}{rn}"
        result.setdefault(pid, {})[k] = {
            'project_id': pid, 'custom_item': ci,
            'section': sec, 'record_name': rn, 'fields_json': fj,
            'company_name': cn or '',
        }
    return result


def _load_existing_cfv(conn: sqlite3.Connection, project_ids: list[int]) -> dict:
    if not project_ids:
        return {}
    placeholders = ','.join('?' * len(project_ids))
    cursor = conn.execute(
        f"SELECT * FROM ContractFieldValues WHERE project_id IN ({placeholders})",
        project_ids,
    )
    cols   = [d[0] for d in cursor.description]
    result: dict[int, dict] = {}
    for row in cursor.fetchall():
        r = dict(zip(cols, row))
        k = f"{r.get('section', '')}{_SEP}{r.get('record_name', '')}"
        result.setdefault(r['project_id'], {})[k] = r
    return result


def _apply_pcf(conn: sqlite3.Connection, to_insert: list[dict], to_delete: list[tuple]):
    for pid, fid in to_delete:
        conn.execute(
            "DELETE FROM ProjectCustomFields WHERE project_id=? AND field_id=?",
            (pid, fid),
        )
    conn.executemany(
        "INSERT INTO ProjectCustomFields (project_id, field_id, field_name, field_type, value) "
        "VALUES (?, ?, ?, ?, ?)",
        [(r['project_id'], r['field_id'], r['field_name'], r['field_type'], r['value'])
         for r in to_insert],
    )


def _apply_contracts(conn: sqlite3.Connection, to_insert: list[dict], to_delete: list[tuple]):
    for pid, ci, sec, rn in to_delete:
        conn.execute(
            "DELETE FROM Contracts WHERE project_id=? AND custom_item=? AND section=? AND record_name=?",
            (pid, ci, sec, rn),
        )
    conn.executemany(
        "INSERT INTO Contracts (project_id, custom_item, section, record_name, fields_json, fetched_at, company_name) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)",
        [(r['project_id'], r['custom_item'], r['section'], r['record_name'], r['fields_json'], r['fetched_at'], r.get('company_name', ''))
         for r in to_insert],
    )


def _apply_cfv(conn: sqlite3.Connection, to_insert: list[dict], to_delete: list[tuple]):
    for pid, sec, rn in to_delete:
        conn.execute(
            "DELETE FROM ContractFieldValues WHERE project_id=? AND section=? AND record_name=?",
            (pid, sec, rn),
        )
    for row in to_insert:
        cols         = list(row.keys())
        col_list     = ', '.join(f'"{c}"' for c in cols)
        placeholders = ', '.join('?' * len(cols))
        conn.execute(
            f"INSERT INTO ContractFieldValues ({col_list}) VALUES ({placeholders})",
            [row[c] for c in cols],
        )


def _update_sync_state(conn: sqlite3.Connection, timestamp_str: str):
    conn.execute(
        "INSERT OR REPLACE INTO SyncState (name, last_sync) VALUES ('Contracts', ?)",
        (timestamp_str,),
    )


# ── Diff functions ────────────────────────────────────────────────────────────

def _diff_pcf(project: dict, existing_by_pid: dict) -> tuple[list[dict], list[tuple]]:
    pid      = project['id']
    existing = existing_by_pid.get(pid, {})

    new_rows = {
        str(cf['id']): {
            'project_id': pid,
            'field_id':   str(cf['id']),
            'field_name': cf.get('name') or '',
            'field_type': cf.get('type') or '',
            'value':      str(cf.get('value') or ''),
        }
        for cf in project.get('customFields', [])
        if cf.get('id')
    }

    to_delete: list[tuple] = []
    to_insert: list[dict]  = []

    for fid, new in new_rows.items():
        old = existing.get(fid)
        if old is None:
            to_insert.append(new)
        elif (old['field_name'] != new['field_name'] or
              old['field_type'] != new['field_type'] or
              old['value']      != new['value']):
            to_delete.append((pid, fid))
            to_insert.append(new)

    for fid in existing:
        if fid not in new_rows:
            to_delete.append((pid, fid))

    return to_insert, to_delete


def _diff_contracts(
    project_id: int, contracts: list[dict],
    existing_by_pid: dict, fetched_at: str, company_name: str = '',
) -> tuple[list[dict], list[tuple]]:
    existing = existing_by_pid.get(project_id, {})

    new_rows: dict[str, dict] = {}
    for c in contracts:
        k = f"{c['customItem']}{_SEP}{c.get('section', '')}{_SEP}{c['record']}"
        new_rows[k] = {
            'project_id':   project_id,
            'custom_item':  c['customItem'],
            'section':      c.get('section', ''),
            'record_name':  c['record'],
            'fields_json':  json.dumps(c['fields']),
            'fetched_at':   fetched_at,
            'company_name': company_name,
        }

    to_delete: list[tuple] = []
    to_insert: list[dict]  = []

    for k, new in new_rows.items():
        old = existing.get(k)
        if old is None:
            to_insert.append(new)
        elif old['fields_json'] != new['fields_json'] or old.get('company_name', '') != new['company_name']:
            ci, sec, rn = k.split(_SEP, 2)
            to_delete.append((project_id, ci, sec, rn))
            to_insert.append(new)

    for k in existing:
        if k not in new_rows:
            ci, sec, rn = k.split(_SEP, 2)
            to_delete.append((project_id, ci, sec, rn))

    return to_insert, to_delete


def _diff_cfv(
    project_id: int, contracts: list[dict],
    existing_by_pid: dict, field_col_map: dict, fetched_at: str,
    company_name: str = '', project_name: str = '',
) -> tuple[list[dict], list[tuple]]:
    wide_records = [c for c in contracts if c['customItem'] == 'AMS Contracts']
    if not wide_records:
        return [], []

    SKIP     = {'fetched_at', 'project_id'}
    existing = existing_by_pid.get(project_id, {})

    new_rows: dict[str, dict] = {}
    for rec in wide_records:
        k   = f"{rec.get('section', '')}{_SEP}{rec['record']}"
        row = {
            'project_id':   project_id,
            'project_name': project_name,
            'section':      rec.get('section', ''),
            'record_name':  rec['record'],
            'fetched_at':   fetched_at,
            'company_name': company_name,
        }
        for raw_name, value in rec['fields'].items():
            row[field_col_map[raw_name]] = str(value) if value is not None else None
        new_rows[k] = row

    to_delete: list[tuple] = []
    to_insert: list[dict]  = []

    for k, new in new_rows.items():
        old = existing.get(k)
        if old is None:
            to_insert.append(new)
        else:
            all_cols = (set(new) | set(old)) - SKIP
            if any(str(new.get(c) or '') != str(old.get(c) or '') for c in all_cols):
                sec, rn = k.split(_SEP, 1)
                to_delete.append((project_id, sec, rn))
                to_insert.append(new)

    for k in existing:
        if k not in new_rows:
            sec, rn = k.split(_SEP, 1)
            to_delete.append((project_id, sec, rn))

    return to_insert, to_delete


# ── Main ──────────────────────────────────────────────────────────────────────

def start_function():
    start_timestamp = datetime.now()
    print(f"Script Started: {start_timestamp}")

    conn = sqlite3.connect(DB_FILE)
    init_tables(conn)
    print(f"SQLite ready: {DB_FILE}\n")

    print("Fetching projects...")
    all_projects = fetch_all_projects()
    print(f"  Total: {len(all_projects)}")

    with_cf = [
        p for p in all_projects
        if 'Application Managed Services' in p.get('name', '') and p.get('customFields')
    ]
    print(f"  Matching with custom fields: {len(with_cf)}\n")

    company_ids = list({p['company']['id'] for p in with_cf if p.get('company', {}).get('id')})
    company_name_map = fetch_company_names(company_ids)
    print(f"  Fetched {len(company_name_map)} company name(s)\n")

    print(f"Fetching contracts ({len(with_cf)} projects, up to {_API_WORKERS} parallel)...")
    all_contracts: dict[int, list[dict]] = {}
    with ThreadPoolExecutor(max_workers=_API_WORKERS) as ex:
        future_to_pid = {ex.submit(fetch_contracts, p['id']): p['id'] for p in with_cf}
        for future in as_completed(future_to_pid):
            pid = future_to_pid[future]
            try:
                all_contracts[pid] = future.result()
            except Exception as e:
                print(f"  Error fetching project {pid}: {e}")
                all_contracts[pid] = []

    active     = [p for p in with_cf if all_contracts.get(p['id'])]
    active_ids = [p['id'] for p in active]
    print(f"  Projects with contracts: {len(active)}\n")

    all_wide = [
        c for pid in active_ids
        for c in all_contracts[pid]
        if c['customItem'] == 'AMS Contracts'
    ]
    raw_fields    = {name for rec in all_wide for name in rec['fields']}
    field_col_map = {name: _col(name) for name in raw_fields}

    _ensure_columns(conn, 'Contracts', {'company_name'})
    cols_cfv = {'company_name', 'project_name'}
    if field_col_map:
        cols_cfv |= set(field_col_map.values())
    _ensure_columns(conn, 'ContractFieldValues', cols_cfv)

    print("Loading existing data from SQLite...")
    existing_pcf       = _load_existing_pcf(conn, active_ids)
    existing_contracts = _load_existing_contracts(conn, active_ids)
    existing_cfv       = _load_existing_cfv(conn, active_ids)

    fetched_at = datetime.now(timezone.utc).isoformat()
    pcf_ins: list[dict]  = []
    pcf_del: list[tuple] = []
    c_ins:   list[dict]  = []
    c_del:   list[tuple] = []
    cfv_ins: list[dict]  = []
    cfv_del: list[tuple] = []

    for p in active:
        pid          = p['id']
        contracts    = all_contracts[pid]
        company_name = company_name_map.get((p.get('company') or {}).get('id'), '')

        ins, dlt = _diff_pcf(p, existing_pcf)
        pcf_ins.extend(ins); pcf_del.extend(dlt)

        ins, dlt = _diff_contracts(pid, contracts, existing_contracts, fetched_at, company_name)
        c_ins.extend(ins); c_del.extend(dlt)

        ins, dlt = _diff_cfv(pid, contracts, existing_cfv, field_col_map, fetched_at, company_name, p['name'])
        cfv_ins.extend(ins); cfv_del.extend(dlt)

    print(f"  ProjectCustomFields  : +{len(pcf_ins)} / -{len(pcf_del)}")
    print(f"  Contracts            : +{len(c_ins)} / -{len(c_del)}")
    print(f"  ContractFieldValues  : +{len(cfv_ins)} / -{len(cfv_del)}")
    print()

    end_timestamp = datetime.now()
    ts_str = end_timestamp.strftime("%Y-%m-%d %H:%M:%S")

    print("Applying changes to SQLite...")
    _apply_pcf(conn, pcf_ins, pcf_del)
    _apply_contracts(conn, c_ins, c_del)
    _apply_cfv(conn, cfv_ins, cfv_del)
    _update_sync_state(conn, ts_str)
    conn.commit()
    conn.close()

    print("Done.")
    print(f"Script Finished: {end_timestamp}")
    print(f"Run Duration (in seconds): {(end_timestamp - start_timestamp).total_seconds()}")


if __name__ == "__main__":
    start_function()
