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
                # also include attributes if any
                for attr_name, attr_value in subchild.attrib.items():
                    row[f"{child.tag}_{subchild.tag}_{attr_name}"] = attr_value
        else:
            row[child.tag] = (child.text or "").strip()

        # include attributes of the child itself
        for attr_name, attr_value in child.attrib.items():
            row[f"{child.tag}_{attr_name}"] = attr_value

    # attributes of element itself
    for attr_name, attr_value in elem.attrib.items():
        row[f"{elem.tag}_{attr_name}"] = attr_value

    return row

def prefixed_table_name(section_tag):
    """Return table name with MAP_ prefix."""
    return f"MAP_{section_tag}"

def extract_xml_section(cur, section):
    """
    Convert one XML top-level section into a SQLite table.
    Handles MappingMeta properly by flattening top-level + nested elements.
    """
    table_name = prefixed_table_name(section.tag)
    print(f"🛠️  Processing XML section: {section.tag} → table {table_name}")

    rows = []

    # If the section has multiple items, treat each child as a row (like SchemaIn/SchemaOut lists)
    # But if the section itself is a single metadata block (like MappingMeta), treat as one row
    if all(not list(child) or child.tag in ("SchemaIn", "SchemaOut") for child in section):
        # Single-row section
        rows = [flatten_xml_row(section)]
    else:
        # Multi-row section (list of elements)
        rows = [flatten_xml_row(r) for r in section]

    if not rows:
        print(f"⚠️  Skipping empty section: {section.tag}")
        return

    columns = sorted({k for row in rows for k in row.keys()})
    col_defs = ", ".join(f'"{col}" TEXT' for col in columns)

    # Recreate table prefixed with MAP_
    cur.execute(f'DROP TABLE IF EXISTS "{table_name}"')
    cur.execute(f'CREATE TABLE "{table_name}" ({col_defs})')

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

    for section in root:
        try:
            extract_xml_section(cur, section)
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