#!/usr/bin/env python3
"""
MIG_GenerateXtendCode.py
========================
Generates Infor M3 Xtend/CTOS extension artifacts (Groovy transaction source,
MI transaction import JSON, and dynamic-table import JSON) from field
metadata stored in the local SQLite "m3xtend" table.

For every distinct MI program / transaction combination found in that table,
this produces the standard Get / Del / Add / Upd / Lst transaction set plus
the underlying EXT dynamic table definition, and writes each as a JSON file
ready to import into M3 (extension source is embedded base64 + SHA256-hashed
inside the transaction JSON):
    DYNAMICDB_<extension_table>.json
    TRANSACTION-<mi_name>-<action><transaction_name>.json

The "m3xtend" table needs columns: miname, trname, MainTable, DynamicTable,
direction, FieldName, MainPrefix, dataType, length, nrOfDecimals,
description, F2FLDI.

Ported from the repo-root GenerateXtendCode.py to a library used by
MIG_App.py: output goes to an explicit out_dir instead of hard-coded
~/Downloads, the DB is ~/sqlite/doppio.db (override with MIG_SYNC_DB, same as
the rest of this package), and generate_all() returns the files written. The
generation logic itself is unchanged - it still works off module-level state
like the original, so generate_all() holds _GEN_LOCK for the whole run.
"""

import base64
import hashlib
import json
import os
import sqlite3
import threading
import uuid
from datetime import datetime, date
from pathlib import Path

SQLITE_DB_PATH = Path(os.environ.get("MIG_SYNC_DB", str(Path.home() / "sqlite" / "doppio.db")))
DEFAULT_OUTPUT_DIR = Path(__file__).parent.resolve() / "output" / "mig_sync" / "xtend_code"

_GEN_LOCK = threading.Lock()


# ---------------------------------------------------------------------------
# FieldInfo dataclass equivalent
# ---------------------------------------------------------------------------

class FieldInfo:
    def __init__(self, name: str, type_: str, length: int, mandatory: bool,
                 description: str, direction: str, nr_of_decimals: int):
        self.name = name
        self.type = type_
        self.length = length
        self.mandatory = mandatory
        self.description = description
        self.direction = direction
        self.nr_of_decimals = nr_of_decimals

    def to_csv_string(self) -> str:
        return (f'"{self.direction}", "{self.name}", "{self.description}", '
                f'{self.length}, {self.mandatory}", "{self.type}"')


# ---------------------------------------------------------------------------
# Module-level state (mirrors Java static fields)
# ---------------------------------------------------------------------------

code: list[str] = []          # built up as a list, joined at the end
out_dir: Path = DEFAULT_OUTPUT_DIR
written_files: list[Path] = []
extension_table: str = ""
fields_array: list[dict] = []
input_fields: list[FieldInfo] = []
keys_array: list[dict] = []
main_prefix: str = ""
main_table: str = ""
mi_name: str = ""
mi_transaction: str = ""
output_fields: list[FieldInfo] = []



# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _create_field_object(name: str, length: int, nr_of_decimals: int, data_type: str) -> dict:
    return {
        "name": name,
        "length": length,
        "nrOfDecimals": nr_of_decimals,
        "dataType": data_type,
    }


def _create_key_object(name: str, ascending: bool) -> dict:
    return {"name": name, "ascending": ascending}


def _title_case(text: str) -> str:
    return " ".join(word.capitalize() for word in text.split())


def _generate_sha256_hash(input_str: str) -> str:
    return hashlib.sha256(input_str.encode()).hexdigest()


def _save_json_to_file(obj: dict, file_name: str) -> None:
    file_path = out_dir / file_name
    with open(file_path, "w") as f:
        json.dump(obj, f)
    written_files.append(file_path)


def _get_formatted_result(sql: str) -> list[dict]:
    """Execute *sql* against the local SQLite DB and return rows as a list of dicts."""
    with sqlite3.connect(SQLITE_DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute(sql)
        return [dict(row) for row in cursor.fetchall()]


def _action_verb(action: str) -> str:
    """Map action code to a descriptive past-tense verb."""
    mapping = {"Get": "get", "Add": "add", "Upd": "update", "Del": "delete", "Lst": "list"}
    for k, v in mapping.items():
        if action.lower() == k.lower():
            return v
    return action.lower()


def _action_title(action: str) -> str:
    """Map action code to a title-cased verb."""
    mapping = {"Get": "Get", "Add": "Add", "Upd": "Update", "Del": "Delete", "Lst": "List"}
    for k, v in mapping.items():
        if action.lower() == k.lower():
            return v
    return action.capitalize()


# ---------------------------------------------------------------------------
# Code generation
# ---------------------------------------------------------------------------

def generate_code(action: str, transaction_name: str,
                  in_fields: list[FieldInfo], out_fields: list[FieldInfo],
                  tbl_main: str, tbl_main_prefix: str, tbl_extension: str) -> None:
    global code

    record_stamp_prefix = "EX" if tbl_extension.startswith("EXT") else tbl_main_prefix
    formatted_date = date.today().strftime("%Y%m%d")

    code = []

    # --- comment block ---
    code.append("/**\n")
    code.append(" * README\n")
    code.append(f" * This extension is being used to {_action_verb(action)} details in the {tbl_extension} table\n")
    code.append(f" * Name: {mi_name}.{action}{transaction_name}\n")
    code.append(f" * Description: {_action_title(action)} details in {tbl_extension} table\n")
    code.append(f" * Date      Changed By            Description\n")
    code.append(f" * {formatted_date}  EPRONOVOST            Added {action}{transaction_name} transaction to access {tbl_extension} table\n")
    code.append(" */\n\n")

    # --- imports (for add / upd) ---
    if action.lower() == "add":
        code.append("import java.time.Instant\n")
        code.append("import java.time.LocalDate\n")
        code.append("import java.time.LocalDateTime\n")
        code.append("import java.time.format.DateTimeFormatter\n\n")
    if action.lower() == "upd":
        code.append("import java.time.Instant\n")
        code.append("import java.time.LocalDate\n")
        code.append("import java.time.format.DateTimeFormatter\n\n")

    # --- class definition ---
    code.append(f"public class {action}{transaction_name} extends ExtendM3Transaction {{\n")
    code.append("\tprivate final MIAPI mi\n")
    code.append("\tprivate final DatabaseAPI database\n")
    code.append("\tprivate final LoggerAPI logger\n")
    code.append("\tprivate final ProgramAPI program\n")
    code.append("\tprivate final UtilityAPI utility\n\n")

    # --- input field declarations ---
    for field in in_fields:
        code.append(f"\tprivate String i{field.name[2:]}\n")
    if action.lower() == "lst":
        code.append("\tprivate int nrOfKeys\n")
    code.append("\n")

    # --- constructor ---
    code.append(f"\tpublic {action}{transaction_name}(MIAPI mi, DatabaseAPI database, LoggerAPI logger, ProgramAPI program, UtilityAPI utility) {{\n")
    code.append("\t\tthis.mi = mi\n")
    code.append("\t\tthis.database = database\n")
    code.append("\t\tthis.logger = logger\n")
    code.append("\t\tthis.program = program\n")
    code.append("\t\tthis.utility = utility\n")
    code.append("\t}\n\n")

    # --- main method ---
    code.append("\t/**\n\t * Main method\n\t *\n\t * @param\n\t * @return\n\t */\n")
    code.append("\tpublic void main() {\n")

    for field in in_fields:
        short = field.name[2:]
        if short in ("CONO", "DIVI"):
            code.append(
                f'\t\ti{short} = mi.inData.get("{short}").isBlank() ? program.getLDAZD().{short} : mi.inData.get("{short}")\n'
            )
        elif field.type == "N":
            code.append(
                f'\t\ti{short} = mi.inData.get("{short}").isBlank() ? "" : mi.inData.get("{short}")\n'
            )
        else:
            code.append(
                f'\t\ti{short} = mi.inData.get("{short}").isBlank() ? "" : mi.inData.get("{short}")\n'
            )

    # nrOfKeys for Lst
    if action.lower() == "lst":
        code.append("\n\t\tnrOfKeys = 0\n")
        for idx, field in enumerate(in_fields, start=1):
            code.append(f'\t\tnrOfKeys = (!i{field.name[2:]}.isBlank() ? {idx} : nrOfKeys)\n')

    # action-specific body
    if action.lower() == "add" and tbl_main != tbl_extension:
        code.append(f'\n\t\tDBAction action = database.table("{tbl_main}").index("00").build()\n')
        code.append("\t\tDBContainer container = action.getContainer()\n\n")
        for field in out_fields:
            short = field.name[2:]
            if field.type == "N":
                ending = ".toDouble())\n" if field.nr_of_decimals != 0 else ".toInteger())\n"
                code.append(f'\t\tcontainer.set("{tbl_main_prefix}{short}", i{short}{ending}')
            else:
                code.append(f'\t\tcontainer.set("{tbl_main_prefix}{short}", i{short})\n')
        code.append("\n")
        code.append("\t\tif (action.read(container)) {\n")
        code.append(f"\t\t\t{action.lower()}Record()\n")
        code.append("\t\t} else {\n")
        code.append('\t\t\tmi.error("Record does not exist")\n')
        code.append("\t\t\treturn\n")
        code.append("\t\t}\n")
        code.append("\t}\n\n")
    else:
        code.append(f"\n\t\t{action.lower()}Record()\n")
        code.append("\t}\n\n")

    # --- xxxRecord method ---
    code.append("\t/**\n")
    code.append(f"\t * {_action_title(action)} the record in {tbl_extension} table\n")
    code.append("\t *\n\t */\n")
    code.append(f"\tprivate void {action.lower()}Record() {{\n")

    if action.lower() in ("get", "lst"):
        code.append(f'\t\tDBAction action = database.table("{tbl_extension}").index("00").selectAllFields().build()\n')
    else:
        code.append(f'\t\tDBAction action = database.table("{tbl_extension}").index("00").build()\n')
    code.append("\t\tDBContainer container = action.getContainer()\n\n")

    if action.lower() == "upd":
        for field in out_fields:
            short = field.name[2:]
            if field.type == "N":
                ending = ".toDouble())\n" if field.nr_of_decimals != 0 else ".toInteger())\n"
                code.append(f'\t\tcontainer.set("{field.name}", i{short}{ending}')
            else:
                code.append(f'\t\tcontainer.set("{field.name}", i{short})\n')
    else:
        for field in in_fields:
            short = field.name[2:]
            if field.type == "N":
                ending = ".toDouble())}\n" if field.nr_of_decimals != 0 else ".toInteger())}\n"
                code.append(f'\t\tif (!i{short}.isBlank()) {{container.set("{field.name}", i{short}{ending}')
            else:
                code.append(f'\t\tif (!i{short}.isBlank()) {{container.set("{field.name}", i{short})}}\n')

    # action-specific closures / reads
    if action.lower() == "get":
        code.append("\n\t\tif(action.read(container)) {\n")
        for field in out_fields:
            code.append(f'\t\t\tmi.outData.put("{field.name[2:]}", container.get("{field.name}").toString())\n')
        code.append("\t\t\tmi.write()\n")
        code.append("\t\t} else {\n")
        code.append('\t\t\tmi.error("Record does not exist")\n')
        code.append("\t\t\treturn\n")
        code.append("\t\t}\n")
        code.append("\t}\n")
        code.append("}")

    elif action.lower() == "lst":
        code.append("\n\t\taction.readAll(container, nrOfKeys, 10000, recordsClosure)\n")
        code.append("\t}\n\n")
        code.append("\t/**\n")
        code.append(f"\t * Closure for {tbl_extension} {_action_title(action)}\n")
        code.append("\t *\n\t */\n")
        code.append("\tClosure<?> recordsClosure = { DBContainer container ->\n")
        for field in out_fields:
            code.append(f'\t\tmi.outData.put("{field.name[2:]}", container.get("{field.name}").toString())\n')
        code.append("\t\tmi.write()\n")
        code.append("\t}\n")
        code.append("}")

    elif action.lower() == "add":
        code.append(f'\n\t\tcontainer.set("{record_stamp_prefix}RGDT", LocalDate.now().format(DateTimeFormatter.ofPattern("yyyyMMdd")).toInteger())\n')
        code.append(f'\t\tcontainer.set("{record_stamp_prefix}RGTM", LocalDateTime.now().format(DateTimeFormatter.ofPattern("HHmmss")).toInteger())\n')
        code.append(f'\t\tcontainer.set("{record_stamp_prefix}LMDT", LocalDate.now().format(DateTimeFormatter.ofPattern("yyyyMMdd")).toInteger())\n')
        code.append(f'\t\tcontainer.set("{record_stamp_prefix}CHNO", 1)\n')
        code.append(f'\t\tcontainer.set("{record_stamp_prefix}CHID", program.getUser())\n')
        code.append("\n\t\taction.insert(container, recordExists)\n")
        code.append("\t}\n\n")
        code.append("\t/**\n")
        code.append(f"\t * Closure for {tbl_extension} {_action_title(action)}\n")
        code.append("\t *\n\t */\n")
        code.append("\tClosure recordExists = {\n")
        code.append('\t\tmi.error("Record already exists")\n')
        code.append("\t}\n")
        code.append("}\n")

    elif action.lower() == "del":
        code.append("\n\t\tif (!action.readLock(container, deleteCallBack)) {\n")
        code.append('\t\t\tmi.error("Record does not exist")\n')
        code.append("\t\t\treturn\n")
        code.append("\t\t}\n")
        code.append("\t}\n\n")
        code.append("\t/**\n")
        code.append(f"\t * Closure for {tbl_extension} {_action_title(action)}\n")
        code.append("\t *\n\t */\n")
        code.append("\tClosure<?> deleteCallBack = { LockedResult lockedResult ->\n")
        code.append("\t\tlockedResult.delete()\n")
        code.append("\t}\n")
        code.append("}\n")

    elif action.lower() == "upd":
        # build fields that are inputs only (not in output keys)
        output_names = {f.name for f in out_fields}
        only_output_fields = [f for f in in_fields if f.name not in output_names]

        code.append("\n\t\tif (!action.readLock(container, updateCallBack)) {\n")
        code.append('\t\t\tmi.error("Record does not exist")\n')
        code.append("\t\t}\n")
        code.append("\t}\n\n")
        code.append("\t/**\n")
        code.append(f"\t * Closure for {tbl_extension} {_action_title(action)}\n")
        code.append("\t *\n\t */\n")
        code.append("\tClosure <?> updateCallBack = {\n")
        code.append("\t\tLockedResult lockedResult ->\n\n")
        code.append(f'\t\tint changeNumber = lockedResult.get("{record_stamp_prefix}CHNO")\n')
        code.append("\t\tint newChangeNumber = changeNumber + 1\n\n")

        for field in only_output_fields:
            short = field.name[2:]
            if field.type == "N":
                ending = ".toDouble())\n\t\t}\n" if field.nr_of_decimals != 0 else ".toInteger())\n\t\t}\n"
                code.append(f'\t\tif (!i{short}.isBlank()) {{\n')
                code.append(f'\t\t\tif (i{short}.trim() == "?") {{i{short} = "0"}}\n')
                code.append(f'\t\t\tlockedResult.set("{field.name}", i{short}{ending}')
            else:
                code.append(f'\t\tif (!i{short}.isBlank()) {{\n')
                code.append(f'\t\t\tif (i{short}.trim() == "?") {{i{short} = " "}}\n')
                code.append(f'\t\t\tlockedResult.set("{field.name}", i{short})\n\t\t}}\n')

        code.append("\n")
        code.append(f'\t\tlockedResult.set("{record_stamp_prefix}LMDT", LocalDate.now().format(DateTimeFormatter.ofPattern("yyyyMMdd")).toInteger())\n')
        code.append(f'\t\tlockedResult.set("{record_stamp_prefix}CHNO", newChangeNumber)\n')
        code.append(f'\t\tlockedResult.set("{record_stamp_prefix}CHID", program.getUser())\n\n')
        code.append("\t\tlockedResult.update()\n")
        code.append("\t}\n")
        code.append("}\n")


# ---------------------------------------------------------------------------
# JSON generation
# ---------------------------------------------------------------------------

def generate_json(action: str, transaction_name: str,
                  in_fields: list[FieldInfo], out_fields: list[FieldInfo]) -> dict:
    source_uuid = str(uuid.uuid4())
    full_transaction = action + transaction_name

    generate_code(action, transaction_name, in_fields, out_fields, main_table, main_prefix, extension_table)

    generated_code = "".join(code)
    base64_code = base64.b64encode(generated_code.encode()).decode()
    code_hash = _generate_sha256_hash(generated_code)

    # input fields
    input_fields_array = []
    if in_fields:
        for field in in_fields:
            input_fields_array.append({
                "name": field.name[2:],
                "description": field.description,
                "length": field.length,
                "mandatory": False if action.lower() == "lst" else field.mandatory,
                "type": field.type,
            })

    # output fields (not for upd/add/del)
    output_fields_array = []
    if action.lower() not in ("upd", "add", "del") and out_fields:
        for field in out_fields:
            output_fields_array.append({
                "name": field.name[2:],
                "description": field.description,
                "length": field.length,
                "mandatory": False if action.lower() == "lst" else field.mandatory,
                "type": field.type,
            })

    description_map = {
        "add": f"Add record into {extension_table}",
        "lst": f"List records from {extension_table}",
        "get": f"Get record into {extension_table}",
        "del": f"Delete record from {extension_table}",
        "upd": f"Update record in {extension_table}",
    }

    transaction_details = {
        "sourceUuid": source_uuid,
        "name": full_transaction,
        "program": mi_name,
        "description": description_map.get(action.lower(), ""),
        "active": True,
        "multi": action.lower() == "lst",
        "modified": 1694126945402,
        "modifiedBy": "EPRONOVOST",
        "outputFields": output_fields_array,
        "inputFields": input_fields_array,
        "utilities": [],
    }

    result = {
        "programModules": {
            mi_name: {
                "program": mi_name,
                "triggers": {},
                "transactions": {full_transaction: transaction_details},
            }
        },
        "utilities": {},
        "sources": {
            source_uuid: {
                "uuid": source_uuid,
                "updated": 1694127214209,
                "updatedBy": "EPRONOVOST",
                "created": 1693515271840,
                "createdBy": "EPRONOVOST",
                "apiVersion": "0.9",
                "beVersion": "16.0.0.20230621154344.6",
                "language": "GROOVY",
                "codeHash": code_hash,
                "code": base64_code,
            }
        },
    }

    _save_json_to_file(result, f"TRANSACTION-{mi_name}-{full_transaction}.json")
    return result


# ---------------------------------------------------------------------------
# Table generation
# ---------------------------------------------------------------------------

def generate_table() -> None:
    global keys_array, fields_array

    table_uuid = str(uuid.uuid4())

    json_object = {
        "type": "DYNAMIC_TABLE",
        "uuid": table_uuid,
        "tableName": extension_table,
        "active": "1",
        "dbCreated": "1",
        "modified": 1654596402217,
        "modifiedBy": "EPRONOVOST",
    }

    keys_array = []
    get_indexes()

    index_object = {
        "name": "00",
        "keys": keys_array,
        "unique": True,
    }

    fields_array = []
    get_fields()

    # Audit / stamp fields
    for fname, dec, dtype in [
        ("EXLMTS", 0, "DECIMAL"),
        ("EXLMDT", 0, "DECIMAL"),
        ("EXRGDT", 0, "DECIMAL"),
        ("EXRGTM", 0, "DECIMAL"),
        ("EXCHNO", 0, "DECIMAL"),
        ("EXCHID", 0, "CHAR"),
    ]:
        fields_array.append(_create_field_object(fname, 18 if fname == "EXLMTS" else 8 if fname in ("EXLMDT", "EXRGDT", "EXRGTM") else 3 if fname == "EXCHNO" else 10, dec, dtype))

    table_object = {
        "name": extension_table,
        "description": "CUSTOM DATABASE TABLE",
        "indexes": [index_object],
        "fields": fields_array,
    }

    json_object["table"] = table_object
    json_object["description"] = f"Extension table for {main_table}"

    _save_json_to_file(json_object, f"DYNAMICDB_{extension_table}.json")


# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------

def get_fields() -> None:
    global output_fields, input_fields, fields_array

    output_fields = []
    input_fields = []

    sql = f"SELECT direction, FieldName, dataType, length, nrOfDecimals, description FROM m3xtend WHERE MainTable='{main_table}'"
    rows = _get_formatted_result(sql)

    for row in rows:
        direction   = str(row.get("direction",    "")).strip()
        name        = str(row.get("FieldName",    "")).strip()
        data_type   = str(row.get("dataType",     "")).strip()
        length      = int(row.get("length",       0))
        nr_decimals = int(row.get("nrOfDecimals", 0))
        description = str(row.get("description",  "")).strip()

        fields_array.append(_create_field_object(name, length, nr_decimals, data_type))

        # normalise type
        typed = data_type
        if typed.upper() in ("CHAR", "STRING"):
            typed = "A"
        elif typed.upper() == "DECIMAL":
            typed = "N"

        field_info = FieldInfo(name, typed, length, False, _title_case(description), "O", nr_decimals)
        output_fields.append(field_info)

        if direction.upper() == "I":
            short = name[2:]
            mandatory = short.upper() not in ("CONO", "DIVI")
            input_fields.append(FieldInfo(name, typed, length, mandatory, _title_case(description), "I", nr_decimals))


def get_indexes() -> None:
    global keys_array, main_prefix

    sql = f"SELECT FieldName, MainPrefix FROM m3xtend WHERE direction = 'I' AND MainTable='{main_table}'"
    rows = _get_formatted_result(sql)

    for row in rows:
        field_name = str(row.get("FieldName",  "")).strip()
        main_prefix = str(row.get("MainPrefix", "")).strip()
        keys_array.append(_create_key_object(field_name, True))


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# m3xtend definitions export / import  (xlsx round-trip for the web tab)
# ---------------------------------------------------------------------------

M3XTEND_COLUMNS = [
    "miname", "trname", "MainTable", "MainTableName", "DynamicTable", "direction",
    "MainPrefix", "FieldName", "description", "dataType", "length", "nrOfDecimals", "F2FLDI",
]
_INT_COLUMNS = {"length", "nrOfDecimals"}
_REQUIRED_COLUMNS = {"miname", "trname", "MainTable", "DynamicTable", "direction", "FieldName"}


def _ensure_table(conn: sqlite3.Connection) -> None:
    cols = ", ".join(f"{c} INTEGER" if c in _INT_COLUMNS else f"{c} VARCHAR(200)" for c in M3XTEND_COLUMNS)
    conn.execute(f"CREATE TABLE IF NOT EXISTS m3xtend ( {cols} )")


def export_definitions(output_dir: Path = DEFAULT_OUTPUT_DIR) -> dict:
    """Write every m3xtend row to M3Xtend_Definitions_<ts>.xlsx in output_dir."""
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    with sqlite3.connect(SQLITE_DB_PATH) as conn:
        _ensure_table(conn)
        rows = conn.execute(
            f"SELECT {', '.join(M3XTEND_COLUMNS)} FROM m3xtend "
            "ORDER BY miname, trname, direction, F2FLDI").fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "m3xtend"
    ws.append(M3XTEND_COLUMNS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(list(row))
    ws.freeze_panes = "A2"
    for i, col in enumerate(M3XTEND_COLUMNS, start=1):
        width = max([len(col)] + [len(str(r[i - 1] or "")) for r in rows])
        ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 50)

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"M3Xtend_Definitions_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    wb.save(path)
    return {"file": path.name, "rows": len(rows)}


def _read_definition_rows(path: Path) -> list[dict]:
    """Rows from an .xlsx (first sheet, or one named m3xtend) or .csv, keyed by header."""
    path = Path(path)
    if path.suffix.lower() == ".csv":
        import csv
        with open(path, newline="", encoding="utf-8-sig") as f:
            table = [r for r in csv.reader(f)]
    else:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb["m3xtend"] if "m3xtend" in wb.sheetnames else wb.worksheets[0]
        table = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    if not table:
        raise ValueError(f"{path.name} is empty.")

    # Match headers case-insensitively to the canonical column names.
    canon = {c.lower(): c for c in M3XTEND_COLUMNS}
    header = [canon.get(str(h or "").strip().lower()) for h in table[0]]
    missing = _REQUIRED_COLUMNS - set(header)
    if missing:
        raise ValueError(f"{path.name} is missing column(s): {', '.join(sorted(missing))}.")

    rows = []
    for line_no, values in enumerate(table[1:], start=2):
        rec = {c: "" for c in M3XTEND_COLUMNS}
        for col, val in zip(header, values):
            if col:
                rec[col] = "" if val is None else str(val).strip()
        if not any(rec.values()):
            continue
        for c in _INT_COLUMNS:
            try:
                rec[c] = int(float(rec[c] or 0))
            except ValueError:
                raise ValueError(f"{path.name} row {line_no}: {c} '{rec[c]}' is not a number.")
        if not rec["miname"] or not rec["FieldName"]:
            raise ValueError(f"{path.name} row {line_no}: miname and FieldName are required.")
        rec["direction"] = rec["direction"].upper()
        rows.append(rec)
    return rows


def _replace_groups(rows: list[dict]) -> dict:
    """
    Every miname/trname group in rows replaces that group's existing m3xtend
    rows; groups not in rows are left alone. All-or-nothing: one transaction.
    """
    groups = sorted({(r["miname"], r["trname"]) for r in rows})
    with sqlite3.connect(SQLITE_DB_PATH) as conn:
        _ensure_table(conn)
        deleted = 0
        for mi, tr in groups:
            deleted += conn.execute(
                "DELETE FROM m3xtend WHERE miname = ? AND trname = ?", (mi, tr)).rowcount
        conn.executemany(
            f"INSERT INTO m3xtend ({', '.join(M3XTEND_COLUMNS)}) "
            f"VALUES ({', '.join('?' for _ in M3XTEND_COLUMNS)})",
            [[r[c] for c in M3XTEND_COLUMNS] for r in rows])
    return {"rows": len(rows), "replaced": deleted,
            "groups": [f"{mi} / {tr}" for mi, tr in groups]}


def import_definitions(path: Path) -> dict:
    """Load m3xtend-format rows from an .xlsx/.csv (see _replace_groups)."""
    return _replace_groups(_read_definition_rows(path))


# ---------------------------------------------------------------------------
# MDP "Table Column" workbooks  (B1 = table name, header row with Column Name,
# Description, Data Type, Length, Decimals, Edit Code, Indexes)
# ---------------------------------------------------------------------------

_MDP_HEADERS = ("Column Name", "Description", "Data Type", "Length", "Decimals", "Indexes")


def read_mdp_workbook(path: Path) -> dict | None:
    """
    {"table": <Table Name:>, "columns": [{Column Name, Description, ...}]} if
    path is an MDP table-column export, else None.
    """
    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        return None
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            table_name, header, columns = "", None, []
            for values in ws.iter_rows(values_only=True):
                cells = ["" if v is None else str(v).strip() for v in values]
                if header is None:
                    if cells and cells[0].lower() == "table name:":
                        table_name = next((c for c in cells[1:] if c), "")
                    elif "Column Name" in cells and "Indexes" in cells:
                        header = {h: cells.index(h) for h in _MDP_HEADERS if h in cells}
                    continue
                rec = {h: (cells[i] if i < len(cells) else "") for h, i in header.items()}
                if rec.get("Column Name"):
                    columns.append(rec)
            if header is not None and table_name:
                return {"table": table_name.upper(), "columns": columns}
    finally:
        wb.close()
    return None


def fetch_table_name(tenant, table: str) -> tuple[str, str]:
    """
    MNS120MI.Get FILE=<table> -> (TX40, "") on success, ("", error message)
    when M3 has no such file. Auth/HTTP failures raise MIG_Api.M3ApiError.
    """
    import requests
    import MIG_Api
    payload = {"program": "MNS120MI",
               "transactions": [{"transaction": "Get", "record": {"FILE": table}}]}
    with requests.Session() as session:
        data = MIG_Api.post_to_m3(tenant, payload, session)
    result = (data.get("results") or [{}])[0]
    if result.get("errorMessage"):
        return "", f"MNS120MI.Get {table}: {result['errorMessage'].strip()}"
    records = result.get("records") or [{}]
    return str(records[0].get("TX40", "")).strip(), ""


def map_mdp_rows(mdp: dict, miname: str, trname: str, dynamic_table: str,
                 main_table_name: str = "") -> list[dict]:
    """Map MDP table columns to m3xtend rows (one per column)."""
    is_ext = dynamic_table.upper().startswith("EXT")
    rows = []
    for col in mdp["columns"]:
        name = col["Column Name"].strip().upper()
        indexes = [i.strip() for i in col.get("Indexes", "").split(",")]

        def _int(field: str) -> int:
            raw = col.get(field, "").strip()
            try:
                return int(float(raw)) if raw else 0
            except ValueError:
                raise ValueError(f"{name}: {field} '{raw}' is not a number.")

        rows.append({
            "miname": miname,
            "trname": trname,
            "MainTable": mdp["table"],
            "MainTableName": main_table_name,
            "DynamicTable": dynamic_table,
            "direction": "I" if "00" in indexes else "O",
            "MainPrefix": name[:2],
            "FieldName": "EX" + name[-4:] if is_ext else name,
            "description": col.get("Description", ""),
            "dataType": col.get("Data Type", ""),
            "length": _int("Length"),
            "nrOfDecimals": _int("Decimals"),
            "F2FLDI": name,
        })
    return rows


def import_mdp_workbook(path: Path, miname: str, trname: str, dynamic_table: str,
                        tenant=None) -> dict:
    """
    Map an MDP workbook into m3xtend. With a tenant, MainTableName is the
    table's TX40 from MNS120MI.Get; a table M3 doesn't know is imported with
    a blank MainTableName and the reason in "warning".
    """
    miname, trname = (miname or "").strip().upper(), (trname or "").strip()
    dynamic_table = (dynamic_table or "").strip().upper()
    if not (miname and trname and dynamic_table):
        raise ValueError("miname, trname and DynamicTable are all required.")
    mdp = read_mdp_workbook(path)
    if mdp is None:
        raise ValueError(f"{Path(path).name} is not an MDP table-column workbook.")
    if not mdp["columns"]:
        raise ValueError(f"{Path(path).name} has no columns under the Column Name header.")
    main_table_name, warning = fetch_table_name(tenant, mdp["table"]) if tenant else ("", "")
    rows = map_mdp_rows(mdp, miname, trname, dynamic_table, main_table_name)
    return {**_replace_groups(rows), "table": mdp["table"],
            "main_table_name": main_table_name, "warning": warning}


GROUPS_SQL = (
    "SELECT DISTINCT miname, trname, MainTable, DynamicTable "
    "FROM m3xtend "
    "WHERE miname <> '' "
    "ORDER BY miname, trname, direction, F2FLDI"
)


def list_groups() -> list[dict]:
    """One row per miname/trname group in m3xtend, with its field count."""
    groups = []
    seen = set()
    for row in _get_formatted_result(GROUPS_SQL):
        key = tuple(str(row.get(k, "")).strip() for k in ("miname", "trname", "MainTable", "DynamicTable"))
        if key in seen:
            continue
        seen.add(key)
        groups.append({"miname": key[0], "trname": key[1], "MainTable": key[2], "DynamicTable": key[3]})
    with sqlite3.connect(SQLITE_DB_PATH) as conn:
        for g in groups:
            g["fields"] = conn.execute(
                "SELECT COUNT(*) FROM m3xtend WHERE MainTable = ?", (g["MainTable"],)).fetchone()[0]
            g["keys"] = conn.execute(
                "SELECT COUNT(*) FROM m3xtend WHERE MainTable = ? AND direction = 'I'", (g["MainTable"],)).fetchone()[0]
            g["MainTableName"] = conn.execute(
                "SELECT COALESCE(MAX(MainTableName), '') FROM m3xtend WHERE miname = ? AND trname = ?",
                (g["miname"], g["trname"])).fetchone()[0]
    return groups


def generate_all(output_dir: Path = DEFAULT_OUTPUT_DIR, progress=None,
                 selected: list[tuple[str, str]] | None = None) -> dict:
    """
    Generate the dynamic table + Get/Del/Add/Upd/Lst transaction JSON for
    every m3xtend group (or only the (miname, trname) pairs in selected) into
    output_dir. progress(done, total, message) is called once per group.
    Returns {"groups": [...], "files": [names]}.
    """
    global mi_name, mi_transaction, main_table, extension_table, out_dir, written_files

    with _GEN_LOCK:
        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        written_files = []

        groups = list_groups()
        if selected is not None:
            wanted = {(mi, tr) for mi, tr in selected}
            groups = [g for g in groups if (g["miname"], g["trname"]) in wanted]
        results = []
        for i, g in enumerate(groups):
            if progress:
                progress(i, len(groups), f"{g['miname']} / {g['trname']} ({g['DynamicTable']})")
            mi_name, mi_transaction = g["miname"], g["trname"]
            main_table, extension_table = g["MainTable"], g["DynamicTable"]
            start = len(written_files)

            generate_table()
            generate_json("Get", mi_transaction, input_fields, output_fields)
            generate_json("Del", mi_transaction, input_fields, output_fields)
            generate_json("Add", mi_transaction, output_fields, input_fields)
            generate_json("Upd", mi_transaction, output_fields, input_fields)
            generate_json("Lst", mi_transaction, input_fields, output_fields)

            results.append({**g, "files": [p.name for p in written_files[start:]]})
        if progress:
            progress(len(groups), len(groups), "Done")

        return {"groups": results, "files": [p.name for p in written_files],
                "output_dir": str(out_dir)}


def main() -> None:
    result = generate_all(Path.home() / "Downloads")
    for name in result["files"]:
        print(name)


if __name__ == "__main__":
    main()
