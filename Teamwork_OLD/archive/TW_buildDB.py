import re
import sqlite3
import requests
import time
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

LOCAL_TZ = ZoneInfo('America/Los_Angeles')

# --- Configuration ---
HOST = 'https://doppiogroup.teamwork.com'
# Teamwork Basic auth expects base64("email:apikey") — not a plain password
API_KEY = 'ZXJpY0Bkb3BwaW9ncm91cC5jb206WnNlNDVyZFhET1BQSU8wMQ=='  # Replace with actual encoded key
HEADERS = {"Authorization": "Basic " + API_KEY}
DB_FILE = 'teamwork.db'

# --- Database Setup ---
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    # Create table to store time entries
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS time_entries (
            id TEXT PRIMARY KEY,
            projectId TEXT,
            projectName TEXT,
            companyName TEXT,
            taskListName TEXT,
            todoitemname TEXT,
            taskIsSubTask TEXT,
            parentTaskName TEXT,
            date DATE,
            startTime TEXT,
            billableHours REAL,
            nonBillableHours REAL,
            billedHours REAL,
            unbilledHours REAL,
            estimatedTime REAL,
            personid TEXT,
            personfirstname TEXT,
            personlastname TEXT,
            taskId TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS invoices (
            id TEXT PRIMARY KEY,
            number TEXT,
            description TEXT,
            status TEXT,
            projectId TEXT,
            projectName TEXT,
            companyId TEXT,
            companyName TEXT,
            poNumber TEXT,
            fixedCost TEXT,
            currencyCode TEXT,
            displayDate TEXT,
            dateCreated TEXT,
            dateUpdated TEXT,
            exportedDate TEXT,
            createdByUserId TEXT,
            createdByUserFirstname TEXT,
            createdByUserLastname TEXT,
            updatedByUserId TEXT,
            exportedByUserId TEXT,
            exportedByUserFirstname TEXT,
            exportedByUserLastname TEXT,
            editedByUserFirstname TEXT,
            editedByUserLastname TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY,
            name TEXT,
            status TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS links (
            id TEXT PRIMARY KEY,
            projectId TEXT,
            projectName TEXT,
            name TEXT,
            description TEXT,
            code TEXT,
            categoryId TEXT,
            categoryName TEXT,
            provider TEXT,
            private TEXT,
            createdDate TEXT,
            updatedDate TEXT,
            createdByUserId TEXT,
            createdByUserFirstname TEXT,
            createdByUserLastname TEXT,
            updatedByUserId TEXT,
            updatedByUserFirstname TEXT,
            updatedByUserLastname TEXT,
            numberOfComments TEXT,
            openInNewWindow TEXT,
            contract_start TEXT,
            contract_end TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sync_state (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    ''')
    conn.commit()
    conn.close()

def get_last_sync():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        cursor.execute('SELECT value FROM sync_state WHERE key = "last_sync"')
        row = cursor.fetchone()
        return row[0] if row else None
    except sqlite3.OperationalError:
        return None
    finally:
        conn.close()

def set_last_sync(dt_str):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('INSERT OR REPLACE INTO sync_state (key, value) VALUES ("last_sync", ?)', (dt_str,))
    conn.commit()
    conn.close()

# --- Data Fetching & Processing ---
def get_json(url):
    # Exponential backoff on 429: Teamwork rate-limits burst fetches heavily
    for attempt in range(5):
        response = requests.get(f"{HOST}{url}", headers=HEADERS)
        if response.status_code == 429:
            wait = 2 ** attempt
            print(f"  Rate limited, retrying in {wait}s...")
            time.sleep(wait)
            continue
        response.raise_for_status()
        return response.json()
    response.raise_for_status()  # re-raise after all retries exhausted

def get_project_ids():
    # Category IDs 1342, 12445, 12446 are the Managed Services project categories
    url = '/projects/api/v3/projects.json?projectStatuses=all&pageSize=400&includeArchivedProjects=true&projectCategoryIds=1342,12445,12446'
    data = get_json(url)
    return [str(p['id']) for p in data['projects']]

def fetch_time_entries(project_id_set, from_date, to_date):
    """Fetch and filter MS time entries for a date range, handling pagination."""
    entries = []
    page = 0
    while True:
        page += 1
        url = f'/time_entries.json?fromdate={from_date}&todate={to_date}&page={page}&pageSize=500&projectType=all'
        data = get_json(url)
        batch = data.get('time-entries', [])
        if not batch:
            break
        ms_batch = [e for e in batch if e['project-id'] in project_id_set]
        entries.extend(ms_batch)
    return entries

def fetch_time_entries_since(project_id_set, updated_after):
    """Fetch MS time entries modified after updated_after (YYYYMMDD), handling pagination."""
    entries = []
    page = 0
    while True:
        page += 1
        url = f'/time_entries.json?updatedAfterDate={updated_after}&page={page}&pageSize=500&projectType=all'
        data = get_json(url)
        batch = data.get('time-entries', [])
        if not batch:
            break
        ms_batch = [e for e in batch if e['project-id'] in project_id_set]
        entries.extend(ms_batch)
    return entries

def get_row(entry):
    # Replicates original logic for normalization
    raw_date = entry['date']
    if 'T' in raw_date:
        # Newer API responses include full ISO 8601 timestamps in UTC; convert to LA local date
        dt = datetime.fromisoformat(raw_date.replace('Z', '+00:00'))
        date_str = dt.astimezone(LOCAL_TZ).strftime('%Y-%m-%d')
    else:
        # Older entries use compact YYYYMMDD format with no timezone
        date_str = datetime.strptime(raw_date, '%Y%m%d').strftime('%Y-%m-%d')
    
    raw_user_date = entry.get('dateUserPerspective', '')
    if raw_user_date and entry.get('has-start-time') == '1':
        start_time = datetime.fromisoformat(raw_user_date.replace('Z', '+00:00')).strftime('%H:%M')
    else:
        start_time = ''

    hours = float(entry.get('hoursDecimal', 0))
    is_billable = entry.get('isbillable') == '1'
    is_billed = entry.get('isbilled') == '1'

    row = {
        'id': entry['id'],
        'projectId': entry['project-id'],
        'projectName': entry['project-name'],
        'companyName': entry['company-name'],
        'date': date_str,
        'startTime': start_time,
        'taskListName': entry.get('todo-list-name', 'General'),
        'todoitemname': entry.get('todo-item-name'),
        'taskIsSubTask': entry.get('taskIsSubTask'),
        'parentTaskName': entry.get('parentTaskName'),
        'personid': entry['person-id'],
        'personfirstname': entry['person-first-name'],
        'personlastname': entry['person-last-name'],
        'taskId': entry.get('todo-item-id'),
        'billableHours': hours if is_billable else 0,
        'nonBillableHours': 0 if is_billable else hours,
        'billedHours': hours if (is_billable and is_billed) else 0,
        'unbilledHours': hours if (is_billable and not is_billed) else 0,
        'estimatedTime': float(entry.get('taskEstimatedTime', 0)) / 60  # API returns minutes; store as hours
    }
    return row

def get_project_ids_from_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('SELECT DISTINCT projectId FROM time_entries')
    ids = [row[0] for row in cursor.fetchall()]
    conn.close()
    return ids

def fetch_invoices(project_ids, updated_after=None):
    invoices = []
    date_filter = f'&updatedAfterDate={updated_after}' if updated_after else ''
    for project_id in project_ids:
        page = 0
        while True:
            page += 1
            data = get_json(f'/projects/{project_id}/invoices.json?type=all&page={page}&pageSize=500{date_filter}')
            batch = data.get('invoices', [])
            if not batch:
                break
            invoices.extend(batch)
    return invoices

def get_invoice_row(inv):
    return {
        'id': inv['id'],
        'number': inv.get('number'),
        'description': inv.get('description'),
        'status': inv.get('status'),
        'projectId': inv.get('project-id'),
        'projectName': inv.get('project-name'),
        'companyId': inv.get('company-id'),
        'companyName': inv.get('company-name'),
        'poNumber': inv.get('po-number'),
        'fixedCost': inv.get('fixed-cost'),
        'currencyCode': inv.get('currency-code'),
        'displayDate': inv.get('display-date'),
        'dateCreated': inv.get('date-created'),
        'dateUpdated': inv.get('date-updated'),
        'exportedDate': inv.get('exported-date'),
        'createdByUserId': inv.get('created-by-user-id'),
        'createdByUserFirstname': inv.get('created-by-user-firstname'),
        'createdByUserLastname': inv.get('created-by-user-lastname'),
        'updatedByUserId': inv.get('update-by-user-id'),
        'exportedByUserId': inv.get('exported-by-user-id'),
        'exportedByUserFirstname': inv.get('exported-by-user-firstname'),
        'exportedByUserLastname': inv.get('exported-by-user-lastname'),
        'editedByUserFirstname': inv.get('edited-by-user-firstname'),
        'editedByUserLastname': inv.get('edited-by-user-lastname'),
    }

def fetch_links(project_ids):
    links = []
    for project_id in project_ids:
        data = get_json(f'/projects/{project_id}/links.json')
        batch = data.get('project', {}).get('links', [])
        links.extend(batch)
    return links

def parse_contract_dates(name, description):
    # Contract dates are stored in two inconsistent formats across projects:
    #   1. Link name parenthetical: "Some Link (Jan 1, 2025 - Dec 31, 2025)"
    #   2. Link description labeled fields: "Start date: Jan 1 2025\nEnd date: Dec 31 2025"
    MONTH_PAT = r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*'
    DATE_PAT = rf'{MONTH_PAT}\s+\d{{1,2}},?\s+\d{{4}}'
    RANGE_RE = re.compile(rf'({DATE_PAT})\s+(?:-to-|-)\s+({DATE_PAT})', re.IGNORECASE)

    def to_iso(s):
        s = re.sub(r',', '', s.strip())
        s = re.sub(r'\s+', ' ', s)
        s = re.sub(r'\bSept\b', 'Sep', s, flags=re.IGNORECASE)  # normalize non-standard abbreviation
        for fmt in ('%b %d %Y', '%B %d %Y'):
            try:
                return datetime.strptime(s, fmt).strftime('%Y/%m/%d')
            except ValueError:
                pass
        return ''

    # Strategy 1: date range inside parentheses in the link name
    if name:
        for content in re.findall(r'\(([^)]+)\)', name):
            m = RANGE_RE.search(content)
            if m:
                return to_iso(m.group(1)), to_iso(m.group(2))

    # Strategy 2: labeled Start/End date fields in the description
    if description:
        start_m = re.search(rf'Start\s+date\s*:\s*({DATE_PAT})', description, re.IGNORECASE)
        end_m = re.search(rf'End\s+date\s*:\s*({DATE_PAT})', description, re.IGNORECASE)
        start = to_iso(start_m.group(1)) if start_m else ''
        end = to_iso(end_m.group(1)) if end_m else ''
        if start or end:
            return start, end

    return '', ''


def get_link_row(link):
    name = link.get('name')
    description = link.get('description')
    contract_start, contract_end = parse_contract_dates(name, description)
    return {
        'id': link['id'],
        'projectId': link.get('project-id'),
        'projectName': link.get('projectName'),
        'name': name,
        'description': description,
        'code': link.get('code'),
        'categoryId': link.get('category-id'),
        'categoryName': link.get('category-name'),
        'provider': link.get('provider'),
        'private': link.get('private'),
        'createdDate': link.get('created-date'),
        'updatedDate': link.get('updated-date'),
        'createdByUserId': link.get('created-by-userId'),
        'createdByUserFirstname': link.get('created-by-userfirstname'),
        'createdByUserLastname': link.get('created-by-userlastname'),
        'updatedByUserId': link.get('updated-by-userId'),
        'updatedByUserFirstname': link.get('updated-by-userfirstname'),
        'updatedByUserLastname': link.get('updated-by-userlastname'),
        'numberOfComments': link.get('numberOfComments'),
        'openInNewWindow': link.get('open-in-new-window'),
        'contract_start': contract_start,
        'contract_end': contract_end,
    }

def save_links_to_db(rows):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    query = '''INSERT OR REPLACE INTO links VALUES
               (:id, :projectId, :projectName, :name, :description, :code,
                :categoryId, :categoryName, :provider, :private,
                :createdDate, :updatedDate,
                :createdByUserId, :createdByUserFirstname, :createdByUserLastname,
                :updatedByUserId, :updatedByUserFirstname, :updatedByUserLastname,
                :numberOfComments, :openInNewWindow,
                :contract_start, :contract_end)'''
    cursor.executemany(query, rows)
    conn.commit()
    conn.close()

def save_invoices_to_db(rows):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    query = '''INSERT OR REPLACE INTO invoices VALUES
               (:id, :number, :description, :status, :projectId, :projectName,
                :companyId, :companyName, :poNumber, :fixedCost, :currencyCode,
                :displayDate, :dateCreated, :dateUpdated, :exportedDate,
                :createdByUserId, :createdByUserFirstname, :createdByUserLastname,
                :updatedByUserId, :exportedByUserId, :exportedByUserFirstname,
                :exportedByUserLastname, :editedByUserFirstname, :editedByUserLastname)'''
    cursor.executemany(query, rows)
    conn.commit()
    conn.close()

def fetch_projects():
    projects = []
    page = 1
    page_size = 500
    while True:
        data = get_json(f'/projects/api/v3/projects.json?projectStatuses=all&pageSize={page_size}&page={page}&includeArchivedProjects=true')
        batch = data.get('projects', [])
        projects.extend(batch)
        meta = data.get('meta', {}).get('page', {})
        total = meta.get('count', '?')
        print(f"  Fetched {len(projects)} / {total} projects...")
        if not meta.get('hasMore', False):
            break
        page += 1
    return projects

def save_projects_to_db(rows):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.executemany(
        'INSERT OR REPLACE INTO projects VALUES (:id, :name, :status)',
        rows
    )
    conn.commit()
    conn.close()

def save_to_db(entries):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    # Using 'INSERT OR REPLACE' to handle updates to existing time entries
    query = '''INSERT OR REPLACE INTO time_entries VALUES
               (:id, :projectId, :projectName, :companyName, :taskListName,
                :todoitemname, :taskIsSubTask, :parentTaskName, :date, :startTime,
                :billableHours, :nonBillableHours, :billedHours, :unbilledHours,
                :estimatedTime, :personid, :personfirstname, :personlastname, :taskId)'''
    
    cursor.executemany(query, entries)
    conn.commit()
    conn.close()

# --- Main Logic ---
def run_sync():
    # Builds/updates a local SQLite snapshot of Teamwork data for Managed Services projects.
    # On first run: full historical sync (time entries 2024-2026 in half-year chunks).
    # On subsequent runs: incremental — only fetches records updated since last sync.
    # All tables use INSERT OR REPLACE so re-runs are safe and idempotent.
    init_db()
    last_sync = get_last_sync()
    today = datetime.now(LOCAL_TZ).strftime('%Y%m%d')

    if last_sync:
        # 7-day lookback buffer catches entries edited retroactively
        incremental_from = (datetime.strptime(last_sync, '%Y%m%d') - timedelta(days=7)).strftime('%Y%m%d')
        print(f"Incremental sync (last sync: {last_sync}, fetching updates from {incremental_from})...")
    else:
        print("First run — performing full historical sync...")

    print("Fetching projects...")
    raw_projects = fetch_projects()
    project_rows = [{'id': p['id'], 'name': p['name'], 'status': p['status']} for p in raw_projects]
    save_projects_to_db(project_rows)
    print(f"  Saved {len(project_rows)} projects.")

    project_id_set = set(get_project_ids())
    print(f"Loaded {len(project_id_set)} MS project IDs.")

    if last_sync:
        print(f"Fetching time entries updated after {incremental_from}...")
        raw = fetch_time_entries_since(project_id_set, incremental_from)
        if raw:
            rows = [get_row(e) for e in raw]
            save_to_db(rows)
            print(f"  Saved {len(rows)} updated entries.")
        else:
            print("  No updated entries found.")
    else:
        for year in range(2024, 2027):
            raw = []
            # Split into half-year chunks to stay within Teamwork's response size limits
            for half, (start, end) in enumerate([('0101', '0630'), ('0701', '1231')], 1):
                label = f"{year} H{half}"
                print(f"Fetching {label}...")
                batch = fetch_time_entries(project_id_set, f'{year}{start}', f'{year}{end}')
                print(f"  {len(batch)} entries")
                raw.extend(batch)
            if raw:
                rows = [get_row(e) for e in raw]
                save_to_db(rows)
                print(f"  Saved {len(rows)} entries for {year}.")
            else:
                print(f"  No entries for {year}.")

    project_ids = get_project_ids_from_db()

    print("Fetching invoices...")
    print(f"  Pulling invoices for {len(project_ids)} projects...")
    raw_invoices = fetch_invoices(project_ids, updated_after=incremental_from if last_sync else None)
    if raw_invoices:
        invoice_rows = [get_invoice_row(i) for i in raw_invoices]
        save_invoices_to_db(invoice_rows)
        print(f"  Saved {len(invoice_rows)} invoices.")
    else:
        print("  No invoices found.")

    print("Fetching links...")
    print(f"  Pulling links for {len(project_ids)} projects...")
    raw_links = fetch_links(project_ids)
    if raw_links:
        link_rows = [get_link_row(l) for l in raw_links]
        save_links_to_db(link_rows)
        print(f"  Saved {len(link_rows)} links.")
    else:
        print("  No links found.")

    set_last_sync(today)
    print("Sync completed.")

if __name__ == "__main__":
    run_sync()