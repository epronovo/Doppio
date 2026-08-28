"""
Fetch all Teamwork "Application Managed Services" projects that have custom fields,
and list their contracts (custom item records).

Outputs a JSON file: projects_with_custom_fields.json
Saves to SQLite: teamwork.db  (tables: project_custom_fields, contracts, contract_field_values)
"""

import json
import sqlite3
import time
from datetime import datetime

import requests

HOST    = 'https://doppiogroup.teamwork.com'
API_KEY = 'ZXJpY0Bkb3BwaW9ncm91cC5jb206WnNlNDVyZFhET1BQSU8wMQ=='
HEADERS = {"Authorization": "Basic " + API_KEY}
DB_FILE = 'teamwork.db'


def get_json(url):
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


def fetch_all_projects():
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


def fetch_contracts(project_id: int) -> list[dict]:
    """Return all contract records for a project, with field names resolved."""
    # Step 1: find custom items on this project
    items_data = get_json(f'/projects/api/v3/projects/{project_id}/customitems.json')
    custom_items = items_data.get('customItems', [])
    if not custom_items:
        return []

    contracts = []
    for item in custom_items:
        item_id = item['id']

        # Step 2: get field definitions (twId uuid → display name)
        fields_data = get_json(f'/projects/api/v3/customitems/{item_id}/fields.json')
        field_map = {
            f['twId']: f['displayName']
            for f in fields_data.get('customItemFields', [])
            if f.get('twId')
        }

        # Step 3: get section names (id → displayName)
        sections_data = get_json(f'/projects/api/v3/customitems/{item_id}/sections.json')
        section_map = {
            s['id']: s['displayName']
            for s in sections_data.get('customItemSections', [])
        }

        # Step 4: get records
        records_data = get_json(f'/projects/api/v3/customitems/{item_id}/records.json')
        for rec in records_data.get('customItemRecords', []):
            named_values = {
                field_map.get(uuid, uuid): value
                for uuid, value in (rec.get('fieldValues') or {}).items()
            }
            section_id   = (rec.get('section') or {}).get('id')
            section_name = section_map.get(section_id, '')
            contracts.append({
                'record':      rec['name'],
                'section':     section_name,
                'customItem':  item.get('displayName', str(item_id)),
                'fields':      named_values,
            })

    return contracts


def init_contract_tables():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS project_custom_fields (
            project_id INTEGER,
            field_id   TEXT,
            field_name TEXT,
            field_type TEXT,
            value      TEXT,
            PRIMARY KEY (project_id, field_id)
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS contracts (
            project_id  INTEGER,
            custom_item TEXT,
            section     TEXT,
            record_name TEXT,
            fields_json TEXT,
            fetched_at  TEXT,
            PRIMARY KEY (project_id, custom_item, section, record_name)
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS contract_field_values (
            project_id  INTEGER,
            section     TEXT,
            record_name TEXT,
            fetched_at  TEXT,
            PRIMARY KEY (project_id, section, record_name)
        )
    ''')
    conn.commit()
    conn.close()


def save_project_to_db(project: dict):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    pid = project['id']
    c.execute(
        'INSERT OR REPLACE INTO projects VALUES (?, ?, ?)',
        (pid, project.get('name'), project.get('status'))
    )
    for cf in project.get('customFields', []):
        c.execute(
            'INSERT OR REPLACE INTO project_custom_fields VALUES (?, ?, ?, ?, ?)',
            (pid, cf.get('id'), cf.get('name'), cf.get('type'), cf.get('value'))
        )
    conn.commit()
    conn.close()


def _ensure_contract_columns(conn: sqlite3.Connection, field_names: set[str]):
    """Add any field columns that don't yet exist in contract_field_values."""
    c = conn.cursor()
    c.execute("PRAGMA table_info(contract_field_values)")
    existing = {row[1] for row in c.fetchall()}
    for name in field_names:
        if name not in existing:
            c.execute(f'ALTER TABLE contract_field_values ADD COLUMN "{name}" TEXT')


def save_contracts_to_db(project_id: int, contracts: list[dict]):
    if not contracts:
        return
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    fetched_at = datetime.now().isoformat()

    for contract in contracts:
        custom_item = contract['customItem']
        section     = contract.get('section', '')
        record_name = contract['record']
        c.execute(
            'INSERT OR REPLACE INTO contracts VALUES (?, ?, ?, ?, ?, ?)',
            (project_id, custom_item, section, record_name,
             json.dumps(contract['fields']), fetched_at)
        )

    # Wide-format table: one row per "Contracts" record, one column per field
    wide_records = [ct for ct in contracts if ct['customItem'] == 'AMS Contracts']
    if wide_records:
        all_fields = {name for rec in wide_records for name in rec['fields']}
        _ensure_contract_columns(conn, all_fields)
        for rec in wide_records:
            fields      = rec['fields']
            field_names = list(fields.keys())
            cols        = ['project_id', 'section', 'record_name', 'fetched_at'] + [f'"{f}"' for f in field_names]
            vals        = [project_id, rec.get('section', ''), rec['record'], fetched_at] + [fields[f] for f in field_names]
            col_str     = ', '.join(cols)
            placeholders = ', '.join(['?'] * len(vals))
            c.execute(
                f'INSERT OR REPLACE INTO contract_field_values ({col_str}) VALUES ({placeholders})',
                vals
            )

    conn.commit()
    conn.close()


def print_contracts(contracts: list[dict]):
    if not contracts:
        print("    (no contracts)")
        return
    for c in contracts:
        section = f"  {c['section']}" if c.get('section') else ''
        print(f"    [{c['customItem']}]{section} — {c['record']}")
        for field_name, value in c['fields'].items():
            if value:
                print(f"      {field_name:<20} {value}")


def main():
    init_contract_tables()

    print("Fetching projects...")
    all_projects = fetch_all_projects()
    print(f"  Total: {len(all_projects)}")

    with_cf = [
        p for p in all_projects
        if 'Application Managed Services' in p.get('name', '') and p.get('customFields')
    ]
    print(f"  Matching with custom fields: {len(with_cf)}\n")

    output = []
    for p in with_cf:
        pid  = p['id']
        name = p.get('name', '')

        contracts = fetch_contracts(pid)
        if not contracts:
            continue

        print(f"{'=' * 70}")
        print(f"  {name}  (id: {pid})")
        print(f"{'=' * 70}")

        print("  Custom Fields:")
        for cf in p['customFields']:
            print(f"    {str(cf.get('name') or ''):<35} {str(cf.get('type') or ''):<15} {cf.get('value') or ''}")

        print("  Contracts:")
        print_contracts(contracts)
        print()

        save_project_to_db(p)
        save_contracts_to_db(pid, contracts)

        p['contracts'] = contracts
        output.append(p)

    out_path = 'projects_with_custom_fields.json'
    with open(out_path, 'w') as f:
        json.dump(output, f, indent=2)
    print(f"Saved → {out_path}")


if __name__ == "__main__":
    main()
