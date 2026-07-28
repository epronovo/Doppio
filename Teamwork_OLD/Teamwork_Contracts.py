"""
Teamwork_Contracts.py
---------------------
Fetch all Teamwork "Application Managed Services" projects with custom fields
and save their contracts (custom item records) to BigQuery.

BigQuery: spatial-earth-492100-b7.teamwork
Tables  : ProjectCustomFields, Contracts, ContractFieldValues, SyncState

Only inserts or removes rows that have actually changed.
"""

import json
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone

import requests

# ── Teamwork ──────────────────────────────────────────────────────────────────
HOST    = 'https://doppiogroup.teamwork.com'
API_KEY = 'ZXJpY0Bkb3BwaW9ncm91cC5jb206WnNlNDVyZFhET1BQSU8wMQ=='
HEADERS = {"Authorization": "Basic " + API_KEY}

# ── BigQuery ──────────────────────────────────────────────────────────────────
BQ_PROJECT = 'spatial-earth-492100-b7'
BQ_DATASET = 'teamwork'
KEY_PATH   = '/Users/ericpronovost/Downloads/spatial-earth-492100-b7-cd5d8fb255b4.json'

_SEP         = '\x1f'
_API_WORKERS = 2
_BQ_WORKERS  = 3


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


# ── BigQuery helpers ──────────────────────────────────────────────────────────

def get_bq_client():
    from google.cloud import bigquery
    from google.oauth2 import service_account
    creds = service_account.Credentials.from_service_account_file(
        KEY_PATH,
        scopes=["https://www.googleapis.com/auth/bigquery"],
    )
    return bigquery.Client(project=BQ_PROJECT, credentials=creds)


def init_bq_tables(client):
    from google.cloud import bigquery
    dataset_ref = bigquery.Dataset(f"{BQ_PROJECT}.{BQ_DATASET}")
    dataset_ref.location = "US"
    client.create_dataset(dataset_ref, exists_ok=True)

    tables = {
        "ProjectCustomFields": [
            bigquery.SchemaField("project_id",  "INTEGER"),
            bigquery.SchemaField("field_id",    "STRING"),
            bigquery.SchemaField("field_name",  "STRING"),
            bigquery.SchemaField("field_type",  "STRING"),
            bigquery.SchemaField("value",       "STRING"),
        ],
        "Contracts": [
            bigquery.SchemaField("project_id",   "INTEGER"),
            bigquery.SchemaField("custom_item",  "STRING"),
            bigquery.SchemaField("section",      "STRING"),
            bigquery.SchemaField("record_name",  "STRING"),
            bigquery.SchemaField("fields_json",  "STRING"),
            bigquery.SchemaField("fetched_at",   "TIMESTAMP"),
            bigquery.SchemaField("company_name", "STRING"),
        ],
        "ContractFieldValues": [
            bigquery.SchemaField("project_id",   "INTEGER"),
            bigquery.SchemaField("project_name", "STRING"),
            bigquery.SchemaField("section",      "STRING"),
            bigquery.SchemaField("record_name",  "STRING"),
            bigquery.SchemaField("fetched_at",   "TIMESTAMP"),
            bigquery.SchemaField("company_name", "STRING"),
        ],
        "SyncState": [
            bigquery.SchemaField("name",      "STRING"),
            bigquery.SchemaField("last_sync", "STRING"),
        ],
    }

    for table_name, schema in tables.items():
        table = bigquery.Table(f"{BQ_PROJECT}.{BQ_DATASET}.{table_name}", schema=schema)
        client.create_table(table, exists_ok=True)


def _ensure_bq_columns(client, table_name: str, col_names: set[str]):
    from google.cloud import bigquery
    table_ref = client.dataset(BQ_DATASET).table(table_name)
    table     = client.get_table(table_ref)
    existing  = {f.name for f in table.schema}
    new_fields = [
        bigquery.SchemaField(col, "STRING")
        for col in col_names
        if col not in existing
    ]
    if not new_fields:
        return
    table.schema = table.schema + new_fields
    client.update_table(table, ["schema"])
    print(f"  Added {len(new_fields)} column(s) to {table_name}")


def _sql_str(s) -> str:
    return "'" + str(s).replace("\\", "\\\\").replace("'", "\\'") + "'"


def _fetch_rows(client, query: str) -> list[dict]:
    return [dict(row) for row in client.query(query).result()]


def _load_rows_bq(client, table_id: str, rows: list[dict]):
    from google.cloud import bigquery
    job_config = bigquery.LoadJobConfig(
        write_disposition=bigquery.WriteDisposition.WRITE_APPEND,
        source_format=bigquery.SourceFormat.NEWLINE_DELIMITED_JSON,
    )
    client.load_table_from_json(rows, table_id, job_config=job_config).result()


def _load_existing_pcf_bq(client, project_ids: list[int]) -> dict:
    if not project_ids:
        return {}
    ids_str = ', '.join(str(p) for p in project_ids)
    result: dict[int, dict] = {}
    for row in _fetch_rows(client,
        f"SELECT project_id, field_id, field_name, field_type, value "
        f"FROM `{BQ_PROJECT}.{BQ_DATASET}.ProjectCustomFields` "
        f"WHERE project_id IN ({ids_str})"
    ):
        result.setdefault(row['project_id'], {})[row['field_id']] = row
    return result


def _load_existing_contracts_bq(client, project_ids: list[int]) -> dict:
    if not project_ids:
        return {}
    ids_str = ', '.join(str(p) for p in project_ids)
    result: dict[int, dict] = {}
    for row in _fetch_rows(client,
        f"SELECT project_id, custom_item, section, record_name, fields_json, company_name "
        f"FROM `{BQ_PROJECT}.{BQ_DATASET}.Contracts` "
        f"WHERE project_id IN ({ids_str})"
    ):
        k = f"{row['custom_item']}{_SEP}{row['section']}{_SEP}{row['record_name']}"
        result.setdefault(row['project_id'], {})[k] = row
    return result


def _load_existing_cfv_bq(client, project_ids: list[int]) -> dict:
    if not project_ids:
        return {}
    ids_str = ', '.join(str(p) for p in project_ids)
    result: dict[int, dict] = {}
    for row in _fetch_rows(client,
        f"SELECT * FROM `{BQ_PROJECT}.{BQ_DATASET}.ContractFieldValues` "
        f"WHERE project_id IN ({ids_str})"
    ):
        k = f"{row.get('section', '')}{_SEP}{row.get('record_name', '')}"
        result.setdefault(row['project_id'], {})[k] = row
    return result


def _batch_delete_pcf_bq(client, deletes: list[tuple]):
    if not deletes:
        return
    by_pid: dict[int, list] = {}
    for pid, fid in deletes:
        by_pid.setdefault(pid, []).append(fid)
    conditions = [
        f"(project_id = {pid} AND field_id IN ({', '.join(str(int(f)) for f in fids)}))"
        for pid, fids in by_pid.items()
    ]
    client.query(
        f"DELETE FROM `{BQ_PROJECT}.{BQ_DATASET}.ProjectCustomFields` "
        f"WHERE {' OR '.join(conditions)}"
    ).result()


def _batch_delete_contracts_bq(client, deletes: list[tuple]):
    if not deletes:
        return
    conditions = [
        f"(project_id = {pid} AND custom_item = {_sql_str(ci)}"
        f" AND section = {_sql_str(sec)} AND record_name = {_sql_str(rn)})"
        for pid, ci, sec, rn in deletes
    ]
    client.query(
        f"DELETE FROM `{BQ_PROJECT}.{BQ_DATASET}.Contracts` "
        f"WHERE {' OR '.join(conditions)}"
    ).result()


def _batch_delete_cfv_bq(client, deletes: list[tuple]):
    if not deletes:
        return
    conditions = [
        f"(project_id = {pid} AND section = {_sql_str(sec)} AND record_name = {_sql_str(rn)})"
        for pid, sec, rn in deletes
    ]
    client.query(
        f"DELETE FROM `{BQ_PROJECT}.{BQ_DATASET}.ContractFieldValues` "
        f"WHERE {' OR '.join(conditions)}"
    ).result()


def _update_sync_state_bq(client, timestamp_str: str):
    client.query(
        f"DELETE FROM `{BQ_PROJECT}.{BQ_DATASET}.SyncState` WHERE name = 'Contracts'"
    ).result()
    _load_rows_bq(client, f"{BQ_PROJECT}.{BQ_DATASET}.SyncState",
                  [{'name': 'Contracts', 'last_sync': timestamp_str}])


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

    client = get_bq_client()
    init_bq_tables(client)
    print(f"BigQuery ready: {BQ_PROJECT}.{BQ_DATASET}\n")

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

    _ensure_bq_columns(client, 'Contracts', {'company_name'})
    cols_cfv = {'company_name', 'project_name'}
    if field_col_map:
        cols_cfv |= set(field_col_map.values())
    _ensure_bq_columns(client, 'ContractFieldValues', cols_cfv)

    print("Loading existing BQ data...")
    if active_ids:
        with ThreadPoolExecutor(max_workers=_BQ_WORKERS) as ex:
            f_pcf = ex.submit(_load_existing_pcf_bq,      client, active_ids)
            f_c   = ex.submit(_load_existing_contracts_bq, client, active_ids)
            f_cfv = ex.submit(_load_existing_cfv_bq,       client, active_ids)
            existing_pcf       = f_pcf.result()
            existing_contracts = f_c.result()
            existing_cfv       = f_cfv.result()
    else:
        existing_pcf = existing_contracts = existing_cfv = {}

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

    print("Applying changes to BigQuery...")
    with ThreadPoolExecutor(max_workers=_BQ_WORKERS) as ex:
        futures = [
            ex.submit(_batch_delete_pcf_bq,       client, pcf_del),
            ex.submit(_batch_delete_contracts_bq,  client, c_del),
            ex.submit(_batch_delete_cfv_bq,        client, cfv_del),
        ]
        for f in futures:
            f.result()
    with ThreadPoolExecutor(max_workers=_BQ_WORKERS) as ex:
        futures = []
        if pcf_ins:
            futures.append(ex.submit(_load_rows_bq, client,
                f"{BQ_PROJECT}.{BQ_DATASET}.ProjectCustomFields", pcf_ins))
        if c_ins:
            futures.append(ex.submit(_load_rows_bq, client,
                f"{BQ_PROJECT}.{BQ_DATASET}.Contracts", c_ins))
        if cfv_ins:
            futures.append(ex.submit(_load_rows_bq, client,
                f"{BQ_PROJECT}.{BQ_DATASET}.ContractFieldValues", cfv_ins))
        for f in futures:
            f.result()
    _update_sync_state_bq(client, ts_str)

    print("Done.")
    print(f"Script Finished: {end_timestamp}")
    print(f"Run Duration (in seconds): {(end_timestamp - start_timestamp).total_seconds()}")


if __name__ == "__main__":
    start_function()
