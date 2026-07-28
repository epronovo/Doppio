"""
Local_TimeData.py
-----------------
SQLite-only version of Teamwork_TimeData.py for local development.
Writes to teamwork.db — no BigQuery dependency.
"""

import sqlite3
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
import requests

# ── Configuration ─────────────────────────────────────────────────────────────
DB_FILE    = 'teamwork.db'
TABLE_NAME = 'TimeData'

COLUMN_RENAME = {
    'projectId':        'Project_ID',
    'projectName':      'Project_Name',
    'companyName':      'Customer_Name',
    'taskListName':     'Task_List_Name',
    'todoitemname':     'Task_Name',
    'taskIsSubTask':    'Is_subtask__Yes_No_',
    'parentTaskName':   'Parent_Task_Name',
    'date':             'Date',
    'billableHours':    'Billable_Hours',
    'nonBillableHours': 'Non_Billable_Hours',
    'billedHours':      'Billed_Hours',
    'unbilledHours':    'Unbilled_Hours',
    'estimatedTime':    'Task_Estimated_Time',
    'personid':         'Person_ID',
    'personfirstname':  'First_Name',
    'personlastname':   'Last_Name',
    'taskId':           'Task_ID',
}

HOST = 'https://doppiogroup.teamwork.com'
API_KEY = 'ZXJpY0Bkb3BwaW9ncm91cC5jb206WnNlNDVyZFhET1BQSU8wMQ=='
HEADERS = {"Authorization": "Basic " + API_KEY}

MS_CATEGORY_IDS = '1342,12445,12446'
LOCAL_TZ = ZoneInfo('America/Los_Angeles')


# ── API helpers ───────────────────────────────────────────────────────────────

def _get_json(endpoint):
    response = requests.get(f"{HOST}{endpoint}", headers=HEADERS)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"API Error: {response.status_code} - {response.text}")
        return None


def _get_row_v3(timelog, included):
    project_id = timelog.get('projectId')
    project = included.get('projects', {}).get(str(project_id), {})
    project_name = project.get('name', '')

    company_id = project.get('companyId')
    company = included.get('companies', {}).get(str(company_id), {})
    company_name = company.get('name', '')

    user_id = timelog.get('userId')
    user = included.get('users', {}).get(str(user_id), {})
    first_name = user.get('firstName', '')
    last_name = user.get('lastName', '')

    task_id = timelog.get('taskId')
    task = included.get('tasks', {}).get(str(task_id), {}) if task_id else {}
    task_name = task.get('name', '')

    tasklist_id = task.get('tasklistId')
    tasklist = included.get('tasklists', {}).get(str(tasklist_id), {}) if tasklist_id else {}
    tasklist_name = tasklist.get('name', '') or 'General'

    parent_task_id = task.get('parentTaskId')
    parent_task = included.get('tasks', {}).get(str(parent_task_id), {}) if parent_task_id else {}
    parent_task_name = parent_task.get('name', '')
    is_subtask = bool(parent_task_id)

    time_logged = timelog.get('timeLogged', '') or ''
    try:
        if 'T' in time_logged:
            dt = datetime.fromisoformat(time_logged.replace('Z', '+00:00'))
        else:
            dt = datetime.fromisoformat(time_logged[:10] + 'T00:00:00+00:00')
        date_string = dt.astimezone(LOCAL_TZ).strftime('%Y-%m-%d')
    except Exception:
        date_string = datetime.now(LOCAL_TZ).strftime('%Y-%m-%d')

    minutes = float(timelog.get('minutes', 0) or 0)
    hours = minutes / 60.0

    is_billable = bool(timelog.get('billable', False))
    is_billed = bool(timelog.get('projectBillingInvoiceId'))

    return {
        'id': timelog.get('id'),
        'projectId': project_id,
        'projectName': project_name,
        'companyName': company_name,
        'taskListName': tasklist_name,
        'date': date_string,
        'billedHours': hours if (is_billable and is_billed) else 0.0,
        'unbilledHours': hours if (is_billable and not is_billed) else 0.0,
        'billableHours': hours if is_billable else 0.0,
        'nonBillableHours': hours if not is_billable else 0.0,
        'personid': user_id,
        'personfirstname': first_name,
        'personlastname': last_name,
        'parentTaskName': parent_task_name,
        'taskIsSubTask': is_subtask,
        'todoitemname': task_name,
        'taskId': task_id,
        'estimatedTime': 0.0,
    }


def _get_entries_time_v3(start_date=None, end_date=None, updated_after=None):
    raw_timelogs = []
    all_included = {
        'projects': {},
        'companies': {},
        'users': {},
        'tasks': {},
        'tasklists': {},
    }

    includes = 'projects,projects.companies,tasks,tasks.tasklists,tasks.parentTasks,users'
    page = 0

    while True:
        page += 1
        params = [
            f'pageSize=500',
            f'page={page}',
            f'projectCategoryIds={MS_CATEGORY_IDS}',
            f'include={includes}',
        ]
        if start_date:
            params.append(f'startDate={start_date}')
        if end_date:
            params.append(f'endDate={end_date}')
        if updated_after:
            params.append(f'updatedAfter={updated_after}')

        url = '/projects/api/v3/time.json?' + '&'.join(params)
        json_data = _get_json(url)

        if not json_data or not json_data.get('timelogs'):
            break

        page_timelogs = json_data['timelogs']
        raw_timelogs.extend(page_timelogs)

        page_included = json_data.get('included', {})
        for key in all_included:
            for entity_id, entity_data in page_included.get(key, {}).items():
                all_included[key][entity_id] = entity_data

        if len(page_timelogs) < 500:
            break

    if not raw_timelogs:
        return pd.DataFrame()

    processed_rows = [_get_row_v3(t, all_included) for t in raw_timelogs]
    df = pd.DataFrame(processed_rows)

    group_cols = ['id', 'date']
    sum_cols = ['billableHours', 'nonBillableHours', 'billedHours', 'unbilledHours']
    other_cols = [col for col in df.columns if col not in sum_cols and col not in group_cols]
    agg_dict = {col: 'first' for col in other_cols}
    for col in sum_cols:
        agg_dict[col] = 'sum'

    df_grouped = df.groupby(group_cols, as_index=False).agg(agg_dict)

    ordered_cols = [
        'id', 'projectId', 'projectName', 'companyName', 'taskListName',
        'todoitemname', 'taskIsSubTask', 'parentTaskName', 'date',
        'billableHours', 'nonBillableHours', 'billedHours', 'unbilledHours',
        'estimatedTime', 'personid', 'personfirstname', 'personlastname', 'taskId'
    ]
    existing_cols = [c for c in ordered_cols if c in df_grouped.columns]

    return df_grouped[existing_cols]


# ── SQLite sync state ─────────────────────────────────────────────────────────

def _get_last_sync():
    try:
        with sqlite3.connect(DB_FILE) as conn:
            result = pd.read_sql("SELECT last_sync FROM SyncState WHERE name = 'TimeData' LIMIT 1", conn)
            if not result.empty and pd.notna(result['last_sync'].iloc[0]):
                return str(result['last_sync'].iloc[0])
    except Exception:
        pass
    return None


def _update_sync_state(conn, timestamp_str):
    conn.execute("CREATE TABLE IF NOT EXISTS SyncState (name TEXT PRIMARY KEY, last_sync TEXT)")
    cursor = conn.execute("UPDATE SyncState SET last_sync = ? WHERE name = 'TimeData'", (timestamp_str,))
    if cursor.rowcount == 0:
        conn.execute("INSERT INTO SyncState (name, last_sync) VALUES ('TimeData', ?)", (timestamp_str,))


# ── Main ──────────────────────────────────────────────────────────────────────

def start_function():
    start_timestamp = datetime.now()
    end_timestamp = start_timestamp
    print(f"Script Started: {start_timestamp}")

    last_sync = _get_last_sync()

    updated_after = None
    if last_sync:
        try:
            dt = datetime.strptime(last_sync[:19], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            dt = datetime.strptime(last_sync[:10], "%Y-%m-%d")
        updated_after = dt.strftime("%Y-%m-%d")
        print(f"Incremental sync — fetching entries updated on or after: {updated_after}")
    else:
        print("No prior sync found — performing full load.")

    all_dfs = []

    if updated_after:
        df_incremental = _get_entries_time_v3(updated_after=updated_after)
        print(f"Incremental entries fetched: {len(df_incremental)}")
        if not df_incremental.empty:
            all_dfs.append(df_incremental)
    else:
        years = [2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026]
        for year in years:
            print(f"Fetching data for year: {year}")
            df_h1 = _get_entries_time_v3(start_date=f"{year}-01-01", end_date=f"{year}-06-30")
            print(f"  H1 ({year}): {len(df_h1)} entries")
            df_h2 = _get_entries_time_v3(start_date=f"{year}-07-01", end_date=f"{year}-12-31")
            print(f"  H2 ({year}): {len(df_h2)} entries")
            for df in [df_h1, df_h2]:
                if not df.empty:
                    all_dfs.append(df)

    if all_dfs:
        print("Consolidating data and writing to SQLite...")
        new_df = pd.concat(all_dfs, ignore_index=True)
        sqlite_df = new_df.rename(columns=COLUMN_RENAME)

        end_timestamp = datetime.now()
        ts_str = end_timestamp.strftime("%Y-%m-%d %H:%M:%S")

        with sqlite3.connect(DB_FILE) as conn:
            if updated_after:
                try:
                    existing_df = pd.read_sql(f"SELECT * FROM {TABLE_NAME}", conn)
                    updated_ids = set(sqlite_df['id'].astype(str))
                    existing_df = existing_df[~existing_df['id'].astype(str).isin(updated_ids)]
                    master_df = pd.concat([existing_df, sqlite_df], ignore_index=True)
                except Exception:
                    master_df = sqlite_df
            else:
                master_df = sqlite_df
            master_df.to_sql(TABLE_NAME, con=conn, if_exists='replace', index=False)
            _update_sync_state(conn, ts_str)
        print(f"Successfully written {len(master_df)} rows to {DB_FILE}")
    else:
        print("No new data found.")
        end_timestamp = datetime.now()

    print("Completed!")
    print(f"Script Finished: {end_timestamp}")
    print(f"Run Duration (in seconds): {(end_timestamp - start_timestamp).total_seconds()}")


if __name__ == "__main__":
    start_function()
