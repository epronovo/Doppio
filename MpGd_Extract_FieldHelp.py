# MpGd_Extract_FieldHelp.py
# -----------------------------------------------------------------------
# PURPOSE
#   Fetches FieldHelp XML documentation for M3 field names from the Infor
#   cloud help endpoint, parses the structured XML into definition text,
#   alternatives (valid values), and program references, then persists the
#   results to m3FieldHelp.  Fields are sourced from several tables to find
#   any that are still missing help content.  A short (last-4-char) fallback
#   ID is tried when the full field name returns no result.
#
# INPUTS
#   - HerffJones_Mapping_Guide, ExtractRules, m3DataStructures (SQLite)
#     — source of field names that need FieldHelp
#   - M3 /help/GB/FieldHelp/xml/MVX/{field} endpoint (HTTP GET)
#
# OUTPUTS
#   - m3FieldHelp   (SQLite) — FieldHelpID, HeadingID, HeadingText,
#                              Definition, Alternatives, ProgramReferences,
#                              RawXML
#
# DEPENDENCIES
#   - config.py, requests
# -----------------------------------------------------------------------

import logging
import requests
import sqlite3
import xml.etree.ElementTree as ET
from tqdm import tqdm

from config import get_sqlite_db_path

logger = logging.getLogger(__name__)

DB_PATH = get_sqlite_db_path()

# -------------------------------------------------------------------
# Fetch FieldHelp XML
# -------------------------------------------------------------------
def fetch_fieldhelp_xml(field_name: str) -> str:
    url = f"https://m3-cm3xprduse1b.m32.m3.us1.mprd.inforcloudsuite.com/help/GB/FieldHelp/xml/MVX/{field_name}?"
    response = requests.get(url)
    response.raise_for_status()
    return response.text

# -------------------------------------------------------------------
# Parse XML to dict
# -------------------------------------------------------------------
def parse_fieldhelp_xml(xml_string: str) -> dict:
    root = ET.fromstring(xml_string)
    ns = {"fh": "http://schemas.intentia.net/fieldhelp"}

    field_id = root.findtext("./header/fh:FieldHelpID", namespaces=ns)
    heading_id = root.findtext("./header/fh:HeadingID", namespaces=ns)
    heading_text = root.findtext("./header/fh:HeadingText", namespaces=ns)

    paragraphs = root.findall(
        ".//fh:FieldDefinition/fh:TextBlock/fh:Paragraph",
        namespaces=ns
    )

    definition = "\n\n".join(p.text.strip() for p in paragraphs if p.text)

    alternatives = []
    for alt in root.findall(".//fh:Alternative", namespaces=ns):
        value = alt.findtext("fh:Value", namespaces=ns)
        desc = alt.findtext("fh:DescriptionText", namespaces=ns)
        alternatives.append(f"{value}: {desc}")

    program_refs = [p.text for p in root.findall(".//fh:ProgramReference", namespaces=ns)]

    return {
        "FieldHelpID": field_id,
        "HeadingID": heading_id,
        "HeadingText": heading_text,
        "Definition": definition,
        "Alternatives": "\n".join(alternatives),
        "ProgramReferences": ", ".join(program_refs),
        "RawXML": xml_string
    }

# -------------------------------------------------------------------
# Save to SQLite (prevents duplicates)
# -------------------------------------------------------------------
def save_to_sqlite(data: dict):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS m3FieldHelp (
            FieldHelpID TEXT PRIMARY KEY,
            HeadingID TEXT,
            HeadingText TEXT,
            Definition TEXT,
            Alternatives TEXT,
            ProgramReferences TEXT,
            RawXML TEXT
        )
    """)

    cur.execute("""
        INSERT INTO m3FieldHelp (FieldHelpID, HeadingID, HeadingText, Definition, Alternatives, ProgramReferences, RawXML)
        VALUES (:FieldHelpID, :HeadingID, :HeadingText, :Definition, :Alternatives, :ProgramReferences, :RawXML)
        ON CONFLICT(FieldHelpID) DO UPDATE SET
            HeadingID = excluded.HeadingID,
            HeadingText = excluded.HeadingText,
            Definition = excluded.Definition,
            Alternatives = excluded.Alternatives,
            ProgramReferences = excluded.ProgramReferences,
            RawXML = excluded.RawXML;
    """, data)

    conn.commit()
    conn.close()

# -------------------------------------------------------------------
# Get all missing m3FieldHelp entries
# -------------------------------------------------------------------
def get_missing_fields():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    queries = [
        """
        SELECT DISTINCT ColumnName
        FROM DoppioGuide 
        LEFT JOIN m3FieldHelp on FieldHelpID = ColumnName
        WHERE ColumnName <> '' AND FieldHelpID is null
        """,
        """
        SELECT DISTINCT fieldname
        FROM ExtractRules r
        LEFT JOIN m3DataStructures ds on ds.DataStructure = r.DataStructure
        LEFT JOIN m3FieldHelp on FieldHelpID = FieldName
        WHERE r.DataStructure <> '' AND FieldHelpID is null
        """,
        """
        SELECT DISTINCT FieldName
        FROM DoppioGuide 
        LEFT JOIN m3FieldHelp on FieldHelpID = FieldName
        WHERE FieldName <> '' AND FieldHelpID is null
        """,
        """
        SELECT DISTINCT substr(FieldName,3,4) as FieldName
        FROM ExtractRules r
        LEFT JOIN m3DataStructures ds on ds.DataStructure = r.DataStructure
        LEFT JOIN m3FieldHelp on FieldHelpID = substring(FieldName,3,4)
        WHERE r.DataStructure <> '' AND FieldHelpID is null
        """
    ]

    missing = set()
    for sql in queries:
        rows = cur.execute(sql).fetchall()
        missing.update(r[0] for r in rows if r[0])

    conn.close()
    return list(missing)

def has_meaningful_content(parsed: dict) -> bool:
    return any([
        parsed.get("HeadingText"),
        parsed.get("Definition"),
        parsed.get("Alternatives"),
        parsed.get("ProgramReferences"),
    ])

# -------------------------------------------------------------------
# Process all missing fields with progress bar
# -------------------------------------------------------------------
def process_all_missing_fieldhelp():
    missing_fields = get_missing_fields()
    logger.info(f"Found {len(missing_fields)} missing FieldHelp entries.")

    for field_name in tqdm(missing_fields, desc="Fetching FieldHelp", unit="field"):
        try:
            xml_data = fetch_fieldhelp_xml(field_name)
            parsed = parse_fieldhelp_xml(xml_data)

            if has_meaningful_content(parsed):
                save_to_sqlite(parsed)
            else:
                save_to_sqlite({"FieldHelpID": field_name, "HeadingID": "", "HeadingText": "", "Definition": "", "Alternatives": "", "ProgramReferences": "", "RawXML": ""})

        except Exception:
            save_to_sqlite({"FieldHelpID": field_name, "HeadingID": "", "HeadingText": "", "Definition": "", "Alternatives": "", "ProgramReferences": "", "RawXML": ""})

# -------------------------------------------------------------------
# Run as script
# -------------------------------------------------------------------
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S"
    )
    process_all_missing_fieldhelp()