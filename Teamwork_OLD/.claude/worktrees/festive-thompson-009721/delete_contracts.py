"""
Delete all contract records from Teamwork projects whose name contains
'Application Managed Services' and that have custom fields configured.

Usage:
    python delete_contracts.py            # dry run — prints what would be deleted
    python delete_contracts.py --confirm  # actually deletes
"""

import sys
import time
import requests

HOST    = 'https://doppiogroup.teamwork.com'
API_KEY = 'ZXJpY0Bkb3BwaW9ncm91cC5jb206WnNlNDVyZFhET1BQSU8wMQ=='
HEADERS = {"Authorization": "Basic " + API_KEY}

DRY_RUN = '--confirm' not in sys.argv


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


def delete(url):
    for attempt in range(5):
        response = requests.delete(f"{HOST}{url}", headers=HEADERS)
        if response.status_code == 429:
            wait = 2 ** attempt
            print(f"  Rate limited, retrying in {wait}s...")
            time.sleep(wait)
            continue
        response.raise_for_status()
        return
    response.raise_for_status()


def fetch_target_projects():
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

        included       = data.get('included', {})
        raw_cfs        = included.get('customfields', [])
        cf_list        = raw_cfs.values() if isinstance(raw_cfs, dict) else raw_cfs
        cf_defs        = {cf['id']: cf for cf in cf_list}
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
        if not data.get('meta', {}).get('page', {}).get('hasMore', False):
            break
        page += 1

    return [
        p for p in projects
        if 'Application Managed Services' in p.get('name', '') and p.get('customFields')
    ]


def fetch_contract_data(project_id: int) -> list[dict]:
    """Return records, sections, tab, and item id for all custom items on a project."""
    items_data = get_json(f'/projects/api/v3/projects/{project_id}/customitems.json')
    if not items_data.get('customItems'):
        return []

    # Build a name → tab_id map from the project's tabs
    tabs_data = get_json(f'/projects/api/v3/projects/{project_id}/tabs.json')
    tab_by_name = {t['name']: t['id'] for t in tabs_data.get('tabs', []) if t['id']}

    result = []
    for item in items_data.get('customItems', []):
        item_id   = item['id']
        item_name = item.get('displayName', str(item_id))

        records_data = get_json(f'/projects/api/v3/customitems/{item_id}/records.json')
        records = [
            {'record_id': rec['id'], 'record_name': rec.get('name', str(rec['id']))}
            for rec in records_data.get('customItemRecords', [])
        ]

        sections = [
            {'section_id': s['id']}
            for s in item.get('sections', [])
        ]

        tab_id = tab_by_name.get(item_name)

        result.append({
            'item_id':   item_id,
            'item_name': item_name,
            'records':   records,
            'sections':  sections,
            'tab_id':    tab_id,
        })
    return result


def main():
    if DRY_RUN:
        print("DRY RUN — pass --confirm to actually delete\n")
    else:
        print("LIVE DELETE — this is irreversible\n")

    print("Fetching projects...")
    projects = fetch_target_projects()
    print(f"  Found {len(projects)} matching project(s)\n")

    total_deleted = 0

    for p in projects:
        pid  = p['id']
        name = p.get('name', '')

        items = fetch_contract_data(pid)
        if not any(i['records'] or i['sections'] for i in items):
            continue

        print(f"  {name}")

        if not DRY_RUN:
            for item in items:
                item_id = item['item_id']

                for r in item['records']:
                    delete(f"/projects/api/v3/customitems/{item_id}/records/{r['record_id']}.json")
                    total_deleted += 1

                for s in item['sections']:
                    delete(f"/projects/api/v3/customitems/{item_id}/sections/{s['section_id']}.json")
                    total_deleted += 1

                if item.get('tab_id'):
                    delete(f"/projects/api/v3/projects/{pid}/tabs/{item['tab_id']}.json")
                    total_deleted += 1

                delete(f"/projects/api/v3/customitems/{item_id}.json")
                total_deleted += 1

    if DRY_RUN:
        print(f"Dry run complete. Re-run with --confirm to delete.")
    else:
        print(f"Deleted {total_deleted} contract record(s).")


if __name__ == "__main__":
    main()
