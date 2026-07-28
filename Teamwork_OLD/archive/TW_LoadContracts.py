#!/usr/bin/env python3
"""
TW_LoadContracts.py
-------------------
Reads contract data from the local SQLite database and loads it into
Teamwork Projects as custom items, sections, fields, and records.

Usage:
    export TW_API_TOKEN="your_bearer_token_here"
    python3 TW_LoadContracts.py

Flow per project:
  1. Create one Custom Item ("Contracts") per project
  2. Create one Section per Contract_Group
  3. Create four Fields (Start Date, End Date, Rate, Hours) – once per project
  4. Create one Record per Contract_Period_Name (within its section)
  5. Patch the four field values onto each record
"""

import sqlite3
import requests
import sys
import os
import time
from collections import defaultdict

# ── Configuration ──────────────────────────────────────────────────────────────

DB_PATH     = "/Users/ericpronovost/Doppio/Teamwork/teamwork.db"
TW_BASE_URL = "https://doppiogroup.teamwork.com/projects/api/v3"
TW_API_TOKEN = 'ZXJpY0Bkb3BwaW9ncm91cC5jb206WnNlNDVyZFhET1BQSU8wMQ==' 

HEADERS = {
    "Content-Type":  "application/json",
    "Authorization": "Basic " + TW_API_TOKEN,
}

# Field definitions: (display name, API type, key used in tw_ids dict)
FIELD_DEFS = [
    ("Start Date", "date",       "Start_Date"),
    ("End Date",   "date",       "End_Date"),
    ("Rate",       "text-short", "Rate"),
    ("Hours",      "text-short", "Hours"),
    ("URL",        "url",        "URL"),
    ("Additional", "text-short", "Additional"),
    ("Overage",    "text-short", "Overage"),
]

# ── Database ───────────────────────────────────────────────────────────────────

def fetch_contract_details():
    """Return all rows from ContractDetails, ordered for deterministic grouping."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("""
        SELECT c.projectId,c.Project_Name,c.Contract_Group,c.Contract_Period_Name,c.Range_Start,c.Range_End,c.Rate,c.Hours,COALESCE(l.code,'') AS [url],c.Additional,c.Overage
        FROM ContractDetails c
        LEFT JOIN links l ON l.projectId = c.projectId AND date(REPLACE(l.contract_start,'/','-')) >= date(c.Range_Start) AND date(REPLACE(l.contract_start,'/','-')) <= date(c.Range_End)
        LEFT JOIN projects p ON p.Id = c.projectId 
        WHERE status = 'active' 
        GROUP BY c.projectId,c.Project_Name,c.Contract_Group,c.Contract_Period_Name,c.Range_Start,c.Range_End
		ORDER BY
		    c.Project_Name,
		    MAX(c.Range_Start) OVER (PARTITION BY c.projectId, c.Contract_Group) DESC,
		    c.Range_Start ASC
    """)
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def group_contracts(rows):
    """
    Returns a nested structure:
        { projectId: {
            "name": Project_Name,
            "groups": {
                Contract_Group: [ row, ... ]
            }
          }
        }
    """
    projects = {}
    for row in rows:
        pid = row["projectId"]
        grp = row["Contract_Group"]
        if pid not in projects:
            projects[pid] = {"name": row["Project_Name"], "groups": {}}
        if grp not in projects[pid]["groups"]:
            projects[pid]["groups"][grp] = []
        projects[pid]["groups"][grp].append(row)
    return projects

# ── API helpers ────────────────────────────────────────────────────────────────

def _request(method, path, body, _retries=3, _delay=2):
    url = f"{TW_BASE_URL}{path}"
    for attempt in range(_retries):
        resp = requests.request(method, url, headers=HEADERS, json=body)
        if resp.status_code == 404 and attempt < _retries - 1:
            time.sleep(_delay)
            continue
        if not resp.ok:
            print(f"\n  ✗ HTTP {resp.status_code} {method} {url}")
            print(f"    Body sent : {body}")
            print(f"    Response  : {resp.text[:1000]}")
            resp.raise_for_status()
        return resp.json()

def api_get(path):
    return _request("GET", path, None)

def api_post(path, body):
    return _request("POST", path, body)

def api_patch(path, body):
    return _request("PATCH", path, body)

# ── Step 1 ─────────────────────────────────────────────────────────────────────

def create_custom_item(project_id):
    """Create a 'Contracts' custom item on the project. Returns customItemId."""
    data = api_post(
        f"/projects/{project_id}/customitems.json",
        {
            "customItem": {
                "displayName":   "AMS Contracts",
                "labelSingular": "contracts",
                "labelPlural":   "contracts",
            },
        },
    )
    return data["customItem"]["id"]

# ── Step 2 ─────────────────────────────────────────────────────────────────────

import re

def format_group_name(name):
    name = re.sub(r'^\d{4}[-./]\d{2}[-./]\d{2}\s*\|\s*', '', name)
    name = re.sub(r'\bRenewal\b\s*', '', name).strip()
    return name

def create_section(custom_item_id, contract_group):
    """Create a section for a contract group. Returns sectionId."""
    display_name = format_group_name(contract_group)
    data = api_post(
        f"/customitems/{custom_item_id}/sections.json",
        {"customItemSection": {"displayName": display_name}},
    )
    section_id = data["customItemSection"]["id"]

    sections_data = api_get(f"/customitems/{custom_item_id}/sections.json")
    known_ids = [s["id"] for s in sections_data.get("customItemSections", [])]
    if section_id not in known_ids:
        print(f"  ⚠ WARNING: sectionId {section_id} not found after creation (visible: {known_ids})")

    return section_id

# ── Step 3 ─────────────────────────────────────────────────────────────────────

def create_fields(custom_item_id):
    """
    Create the seven standard fields on the custom item.
    Fields live at the customItem level and are shared across all sections/records.
    Called once per project (after the custom item is created).

    Returns a dict: { "Start_Date": twId, "End_Date": twId, "Rate": twId, "Hours": twId, "URL": twId, "Additional": twId, "Overage": twId }
    """
    tw_ids = {}
    for display_name, field_type, key in FIELD_DEFS:
        data = api_post(
            f"/customitems/{custom_item_id}/fields.json",
            {"customItemField": {"displayName": display_name, "type": field_type}},
        )
        tw_ids[key] = data["customItemField"]["twId"]
    return tw_ids

# ── Step 4 ─────────────────────────────────────────────────────────────────────

def create_record(custom_item_id, section_id, period_name):
    """Create a record for a contract period. Returns recordId."""
    data = api_post(
        f"/customitems/{custom_item_id}/records.json",
        {
            "customItemRecord": {
                "name":            period_name,
                "positionAfterId": 0,
                "sectionId":       section_id,
            }
        },
    )
    return data["customItemRecord"]["id"]

# ── Step 5 ─────────────────────────────────────────────────────────────────────

def patch_field_values(custom_item_id, record_id, tw_ids, row):
    """Patch field values onto the record, one PATCH call each. URL is skipped when empty."""
    field_patches = [
        ("Start_Date", "Range_Start"),
        ("End_Date",   "Range_End"),
        ("Rate",       "Rate"),
        ("Hours",      "Hours"),
        ("Additional", "Additional"),
        ("Overage",    "Overage"),
    ]
    for key, db_col in field_patches:
        tw_id = tw_ids[key]
        value = row[db_col]
        str_value = str(value) if value is not None else ""
        if key in ("Start_Date", "End_Date"):
            str_value = str_value.replace(".", "-")
        api_patch(
            f"/customitems/{custom_item_id}/records/{record_id}.json",
            {"customItemRecord": {"fieldValues": {tw_id: str_value}}},
        )

    url_value = row.get("url") or ""
    if url_value:
        tw_id = tw_ids["URL"]
        api_patch(
            f"/customitems/{custom_item_id}/records/{record_id}.json",
            {"customItemRecord": {"fieldValues": {tw_id: f"[Contract]({url_value})"}}},
        )

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    # ── Pre-flight checks ──────────────────────────────────────────────────────
    if not TW_API_TOKEN:
        sys.exit(
            "ERROR: TW_API_TOKEN environment variable is not set.\n"
            "       Export it before running:  export TW_API_TOKEN='your_token'"
        )

    print("=" * 65)
    print("  TW_LoadContracts.py  —  Teamwork Contract Loader")
    print("=" * 65)
    print(f"  Database : {DB_PATH}")
    print(f"  API base : {TW_BASE_URL}")
    print()

    # ── Fetch data ─────────────────────────────────────────────────────────────
    rows = fetch_contract_details()
    if not rows:
        sys.exit("No rows returned from ContractDetails. Nothing to do.")
    print(f"Fetched {len(rows)} contract row(s) from database.\n")

    projects = group_contracts(rows)
    print(f"Found {len(projects)} project(s) to process.\n")

    # ── Process each project ───────────────────────────────────────────────────
    for project_id, project_data in projects.items():
        project_name = project_data["name"]
        groups       = project_data["groups"]

        print(f"  Loading: {project_name}...")

        # Step 1 — One custom item per project
        custom_item_id = create_custom_item(project_id)

        # Step 3 — Fields are at the customItem level (shared across all groups/sections).
        #          Create them once here, immediately after the custom item.
        tw_ids = create_fields(custom_item_id)

        # Steps 2 / 4 / 5 — iterate over groups and their periods
        for contract_group, period_rows in groups.items():

            # Step 2 — One section per (project, group)
            section_id = create_section(custom_item_id, contract_group)

            for row in period_rows:
                period_name = row["Contract_Period_Name"]

                # Step 4 — One record per period
                record_id = create_record(custom_item_id, section_id, period_name)

                # Step 5 — Patch the four field values
                patch_field_values(custom_item_id, record_id, tw_ids, row)

    print("=" * 65)
    print("  Done! All contracts loaded successfully.")
    print("=" * 65)


if __name__ == "__main__":
    main()
