# ExtractXMLMap.py

import os
import time
import sqlite3
import xml.etree.ElementTree as ET

from tqdm import tqdm
from pathlib import Path
from config import BASE_DIR, get_sqlite_db_path
from UserDefaults import load_user_defaults, save_user_defaults
from ExtractAPI import prompt_for_company_division, CONFIG

MAPS_DIR = BASE_DIR / "maps"       # Folder for your .map XML files

# Load defaults
defaults = load_user_defaults()
save_user_defaults(defaults)

# Prompt for company/division and DB name (from ExtractAPI logic)
prompt_for_company_division()

# Fixed folder
DB_FOLDER = Path("/Users/ericpronovost/sqlite")
DB_FOLDER.mkdir(parents=True, exist_ok=True)

# Full SQLite DB path
SQLITE_DB_PATH = DB_FOLDER / CONFIG["local_db_name"]

print(f"📌 Using SQLite DB: {SQLITE_DB_PATH}")


def flatten_xml_row(elem):
    """
    Flatten an XML element into a dict of {column: value}.
    Includes top-level children and nested children like SchemaIn/SchemaOut.
    """
    row = {}

    for child in elem:
        if list(child):  # nested elements
            for subchild in child:
                key = f"{child.tag}_{subchild.tag}"
                row[key] = (subchild.text or "").strip()
                # include attributes of nested elements
                for attr_name, attr_value in subchild.attrib.items():
                    row[f"{child.tag}_{subchild.tag}_{attr_name}"] = attr_value
        else:
            row[child.tag] = (child.text or "").strip()

        # include attributes of the child itself
        for attr_name, attr_value in child.attrib.items():
            row[f"{child.tag}_{attr_name}"] = attr_value

    # include attributes of the element itself
    for attr_name, attr_value in elem.attrib.items():
        row[f"{elem.tag}_{attr_name}"] = attr_value

    return row


def prefixed_table_name(section_tag):
    """Return table name with MAP_ prefix."""
    return f"MAP_{section_tag}"


def ensure_table(cur, table_name, rows):
    """
    Create table if it doesn't exist; add missing columns if needed.
    Ensures MappingMeta_ID is the first column if present.
    """
    cur.execute(f'PRAGMA table_info("{table_name}")')
    existing_cols = [row[1] for row in cur.fetchall()]

    all_cols = sorted({k for row in rows for k in row.keys()})
    
    # Make MappingMeta_ID first if present
    if "MappingMeta_ID" in all_cols:
        all_cols = ["MappingMeta_ID"] + [c for c in all_cols if c != "MappingMeta_ID"]

    missing_cols = [col for col in all_cols if col not in existing_cols]

    if not existing_cols:
        # Table does not exist → create
        col_defs = ", ".join(f'"{col}" TEXT' for col in all_cols)
        cur.execute(f'CREATE TABLE "{table_name}" ({col_defs})')
    elif missing_cols:
        # Table exists → add missing columns at the end
        for col in missing_cols:
            cur.execute(f'ALTER TABLE "{table_name}" ADD COLUMN "{col}" TEXT')


def extract_xml_section(cur, section, mapping_id=None):
    """
    Convert one XML top-level section into a SQLite table.
    If mapping_id is provided, add it to all rows as MappingMeta_ID.
    """
    table_name = prefixed_table_name(section.tag)
    print(f"🛠️  Processing XML section: {section.tag} → table {table_name}")

    rows = []
    if all(not list(child) or child.tag in ("SchemaIn", "SchemaOut") for child in section):
        rows = [flatten_xml_row(section)]
    else:
        rows = [flatten_xml_row(r) for r in section]

    if not rows:
        print(f"⚠️  Skipping empty section: {section.tag}")
        return

    # Add MappingMeta_ID to all rows
    if mapping_id is not None:
        for row in rows:
            row["MappingMeta_ID"] = str(mapping_id)

    # Ensure table exists and columns match
    ensure_table(cur, table_name, rows)

    # Make MappingMeta_ID first in insert order
    columns = sorted({k for row in rows for k in row.keys()})
    if "MappingMeta_ID" in columns:
        columns = ["MappingMeta_ID"] + [c for c in columns if c != "MappingMeta_ID"]

    placeholders = ", ".join(["?"] * len(columns))
    insert_sql = f'INSERT INTO "{table_name}" ({", ".join(columns)}) VALUES ({placeholders})'

    for row in tqdm(rows, desc=f"Inserting {section.tag}"):
        cur.execute(insert_sql, [row.get(col) for col in columns])

    cur.connection.commit()
    print(f"✅ Finished section: {section.tag} → table {table_name} ({len(rows)} rows)")
    

def extract_map_file(cur, file_path):
    print(f"\n📄 Loading map file: {file_path}")
    tree = ET.parse(file_path)
    root = tree.getroot()

    mapping_id = None

    # Extract MappingMeta first to get its ID
    for section in root:
        if section.tag == "MappingMeta":
            try:
                extract_xml_section(cur, section)
                mapping_id = section.findtext("ID")
            except Exception as e:
                print(f"❌ Error processing MappingMeta: {e}")
            break

    # Extract remaining sections with MappingMeta_ID
    for section in root:
        if section.tag != "MappingMeta":
            try:
                extract_xml_section(cur, section, mapping_id=mapping_id)
            except Exception as e:
                print(f"❌ Error processing section {section.tag}: {e}")


def process_map_files():
    return [
        MAPS_DIR / f
        for f in os.listdir(MAPS_DIR)
        if f.endswith(".map") or f.endswith(".xml")
    ]


if __name__ == "__main__":
    print("")
    start = time.time()

    map_files = process_map_files()

    with sqlite3.connect(SQLITE_DB_PATH) as conn:
        cur = conn.cursor()

        for file_path in map_files:
            try:
                extract_map_file(cur, file_path)
            except Exception as e:
                print(f"❌ Error in {file_path}: {e}")

    elapsed = time.time() - start
    h, rem = divmod(elapsed, 3600)
    m, s = divmod(rem, 60)
    print(f"\n✅ Total run time: {int(h):02d}:{int(m):02d}:{int(s):02d} (hh:mm:ss)")