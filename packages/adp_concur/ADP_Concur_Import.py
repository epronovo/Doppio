"""
ADP_Concur_Import - load Kelly's workbook into the database.

One workbook carries everything: the ADP export on a data sheet, six lookup
tabs, and the three Concur record templates. Nothing here depends on the tab
names, because a re-cut of the report will not keep them - each sheet is
recognised by what its header row says, and the header row is found by looking
for it rather than assumed to be row 1. That matters for the ADP sheet, which
has three rows of Kelly's notes above the real headings.

The lookup tabs are a full refresh: dropping a workbook replaces the maps with
what it carries, because the maps are the workbook's job. The employees merge
on File Number, so re-dropping a fresh ADP cut refreshes the people already
held without disturbing the ones keyed in by hand.
"""
from __future__ import annotations

import argparse
import json
import re
import sqlite3
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from ADP_Concur_Db import (
    ADP_COLUMNS,
    DEFAULT_DB_PATH,
    UKG_COLUMNS,
    connect,
    resolve_db_path,
)

# How each sheet is recognised: every one of these headings has to be present
# in the candidate header row. Deliberately short lists - enough to be sure,
# few enough that an extra or renamed column elsewhere does not break the match.
SHEET_SIGNATURES: list[tuple[str, list[str]]] = [
    ("employees",      ["payroll company code", "file number", "position status"]),
    # UKG's export. Three headings no other sheet has together - it is the only
    # one with an Employee Number beside a Salary Grade and a Pay Group.
    ("ukg",            ["employee number", "salary grade", "pay group"]),
    ("status_map",     ["position status", "concur status"]),
    ("country_map",    ["adp country"]),
    ("org_map",        ["business unit description", "home department code"]),
    ("language_map",   ["language", "adp language"]),
    ("salary_map",     ["pay grade code", "expense map"]),
    ("supervisor_map", ["exception employee id", "supervisor id"]),
    ("role_map",       ["role", "assign role automatically"]),
    ("invoice_map",    ["invoice exception employee", "invoice access value"]),
]

# The record templates announce themselves in their first cell.
RECORD_SIGNATURES = {"trx type (305)": "305",
                     "trx type (350)": "350",
                     "trx type (360)": "360",
                     "trx type (700)": "700"}

# Record tabs that are out of scope. The workbook carries them, nothing reads
# them, and they are named in the load report rather than reported as
# unrecognised - "ignored on purpose" and "not understood" are different things
# and a sheet that quietly vanishes is how a record type gets forgotten.
OUT_OF_SCOPE = ("400", "320")

# How far down a sheet to look for the header row.
HEADER_SCAN_ROWS = 8


def _norm(value) -> str:
    """A heading reduced to something comparable - case and whitespace out."""
    if value is None:
        return ""
    return " ".join(str(value).split()).strip().lower()


def _cell(value) -> str:
    """A data cell as the string the database stores."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d") if (value.hour, value.minute, value.second) == (0, 0, 0) \
            else value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def find_header_row(rows: list[tuple], wanted: list[str]) -> int | None:
    """
    The 1-based row that carries every heading in `wanted`.

    Matching is on a leading substring, so 'language' finds 'Language' and
    'supervisor id' finds 'Supervisor ID (from ADP)'. The first row that has
    them all wins.
    """
    for i, row in enumerate(rows[:HEADER_SCAN_ROWS], 1):
        heads = [_norm(c) for c in row]
        if all(any(h.startswith(w) or w in h for h in heads if h) for w in wanted):
            return i
    return None


def identify_sheet(rows: list[tuple], title: str = "") -> tuple[str | None, int | None]:
    """
    What a worksheet is, and which row its headings are on.

    Everything here is decided from the header row - no sheet name is
    consulted, which matters because the 15 September workbook renamed every
    tab. 'ADP' was '1', and the record tabs grew names like '305 ADP Employee
    (All)'. Nothing broke, because nothing was ever looking at the names.

    The one rule that is not a header match is the record tabs, which announce
    themselves in their first cell as 'Trx Type (305)'. The four in scope are
    loaded as layouts; 400 and 320 are recognised and named as skipped.
    """
    for i, row in enumerate(rows[:HEADER_SCAN_ROWS], 1):
        if not row:
            continue
        first = _norm(row[0])
        if first in RECORD_SIGNATURES:
            return "record_" + RECORD_SIGNATURES[first], i
        # A record tab nobody reads. Matched on the same 'Trx Type (nnn)' cell
        # so it is recognised rather than guessed at from the sheet name.
        m = re.fullmatch(r"trx type \((\d{3})\)", first)
        if m and m.group(1) in OUT_OF_SCOPE:
            return "skip_" + m.group(1), i

    for kind, wanted in SHEET_SIGNATURES:
        row = find_header_row(rows, wanted)
        if row:
            return kind, row
    return None, None


def column_index(headings: list[str], *wanted: str) -> int | None:
    """
    Position of the heading that matches any of `wanted`, exact match first.

    Prefix matching is what lets a heading be named loosely - 'Employee ID'
    finds 'Employee  ID (Cannot be changed using this record ...)'. But a
    prefix alone picks the wrong column whenever one heading is the start of
    another, and UKG's sheet is full of those: 'Job' is the start of 'Job
    Code', 'Job Family', 'Job Role' and 'Job Type', and 'Pay Group' is the
    start of 'Pay Group Code'. Asking for 'Job' used to return the job code.

    So each term gets an exact scan before a prefix scan - but one term at a
    time, in the order the caller listed them. The order of `wanted` is the
    caller saying which heading it would rather have, and that has to win over
    exactness: the Country Map asks for 'Legal / Preferred Address' before
    'ADP Country', and a global exact-first pass would hand back the ADP
    Country column, which on that sheet is the value rather than the key.
    """
    normed = [_norm(h) for h in headings]
    for w in wanted:
        w = _norm(w)
        for i, h in enumerate(normed):
            if h == w:
                return i
        for i, h in enumerate(normed):
            if h.startswith(w):
                return i
    return None


# --------------------------------------------------------------- the sheets


# Position Status values that mean the person is still employed. Used only to
# pick between two rows for the same person - what Concur is told comes from
# the Status Map, not from here.
LIVE_STATUSES = {"active", "leave", "leave of absence"}


def _row_rank(row: dict) -> tuple:
    """
    How good a candidate one ADP row is for being *the* row for a person.

    ADP writes one row per employment record, so an internal transfer arrives
    as the same File Number twice - terminated under the old payroll company
    and active under the new one. Concur wants one profile per employee, and
    it wants the live one. Ranking, most significant first:

      1. a live Position Status beats a terminated one
      2. no Termination Date beats having one
      3. the later Rehire Date, then the later Hire Date
      4. having a Supervisor ID beats not having one
      5. the later row in the file

    Every rejected row is recorded on the employee, so the choice is visible
    rather than silent.
    """
    status = (row.get("position_status") or "").strip().lower()
    return (
        1 if status in LIVE_STATUSES else 0,
        0 if (row.get("termination_date") or "").strip() else 1,
        (row.get("rehire_date") or "").strip(),
        (row.get("hire_date") or "").strip(),
        1 if (row.get("supervisor_id_raw") or "").strip() else 0,
        row["__row"],
    )


def _protected_set_clause(columns: list[str]) -> str:
    """
    An `ON CONFLICT ... DO UPDATE SET` clause that leaves a column alone when
    its name is in the row's own `overridden_fields` - a hand edit on the
    Employees tab (see api_employee_save() in ADP_Concur_App.py) - and takes
    the workbook's value otherwise.

    An unqualified column name inside the SET clause of an upsert refers to
    the row as it was *before* this statement, so `col` below is the value
    already on file and `excluded.col` is the one this import would
    otherwise write. `overridden_fields` is comma-bounded on both sides so a
    field name is never mistaken for a substring of another one.
    """
    return ", ".join(
        f"{c} = CASE WHEN (',' || COALESCE(overridden_fields, '') || ',') "
        f"LIKE '%,{c},%' THEN {c} ELSE excluded.{c} END"
        for c in columns if c != "file_number")


def import_employees(conn: sqlite3.Connection, rows: list[tuple], header_row: int,
                     import_id: int) -> dict:
    """
    Merge the ADP export into ADP_Concur_Employees on File Number.

    Only the ADP columns are written. The derived ones are left to
    ADP_Concur_Map.ADP_Concur_derive(), which the caller runs afterwards -
    a workbook that carries stale lookup results should not put them in the
    database.

    A row already held keeps its key, its include flags, its manual edits to
    fields the file does not carry, and - via _protected_set_clause() - any
    field a hand edit deliberately changed away from what ADP sent; a person
    the file does not mention is left alone rather than deleted, because a
    partial cut of ADP is a normal thing to be handed.
    """
    headings = [h for h in rows[header_row - 1]]
    positions = {}
    missing = []
    for heading, column in ADP_COLUMNS:
        idx = column_index(headings, heading)
        if idx is None:
            missing.append(heading)
        positions[column] = idx

    if positions.get("file_number") is None:
        raise ValueError("The ADP sheet has no 'File Number' column - that is "
                         "the key every record is built on.")

    cols = [c for c in positions if positions[c] is not None]

    # Read the whole sheet first, then resolve the duplicates, then write. It
    # has to be done in that order: which row wins is a property of the group,
    # not of the row, so there is nothing to insert until the group is known.
    by_file: dict[str, list[dict]] = {}
    skipped = 0
    for n, row in enumerate(rows[header_row:], header_row + 1):
        file_number = _cell(row[positions["file_number"]]
                            if positions["file_number"] < len(row) else None)
        if not file_number:
            # A wholly blank row is the end of the data, not an error; a row
            # with data but no file number is one we cannot key.
            if any(_cell(c) for c in row):
                skipped += 1
            continue
        record = {"__row": n}
        for c in cols:
            idx = positions[c]
            record[c] = _cell(row[idx]) if idx < len(row) else ""
        record["file_number"] = file_number
        by_file.setdefault(file_number, []).append(record)

    insert_sql = (
        "INSERT INTO ADP_Concur_Employees (source, import_id, source_row, "
        "duplicate_rows, duplicate_note, " + ", ".join(cols) + ") "
        "VALUES ('adp', ?, ?, ?, ?, " + ", ".join("?" for _ in cols) + ") "
        "ON CONFLICT (file_number) DO UPDATE SET "
        + _protected_set_clause(cols)
        + ", import_id = excluded.import_id, source_row = excluded.source_row"
        + ", duplicate_rows = excluded.duplicate_rows"
        + ", duplicate_note = excluded.duplicate_note"
        + ", modified_at = datetime('now')"
        + ", row_state = CASE WHEN ADP_Concur_Employees.row_state = 'new' "
          "THEN 'new' ELSE 'unchanged' END"
    )

    cur = conn.cursor()
    loaded = 0
    duplicates = 0
    for file_number, group in by_file.items():
        note = None
        if len(group) > 1:
            duplicates += 1
            group = sorted(group, key=_row_rank, reverse=True)
            taken, rejected = group[0], group[1:]
            note = (f"{len(group)} ADP rows. Took row {taken['__row']} "
                    f"({taken.get('payroll_company_code') or '?'}, "
                    f"{taken.get('position_status') or '?'}); ignored "
                    + ", ".join(f"row {r['__row']} ({r.get('payroll_company_code') or '?'}, "
                                f"{r.get('position_status') or '?'})" for r in rejected)
                    + ".")
        else:
            taken = group[0]
        cur.execute(insert_sql,
                    [import_id, taken["__row"], len(group), note]
                    + [taken[c] for c in cols])
        loaded += 1

    return {"loaded": loaded, "skipped": skipped, "duplicates": duplicates,
            "source_rows": sum(len(g) for g in by_file.values()),
            "missing_columns": missing}


def _refresh(conn: sqlite3.Connection, table: str, columns: list[str],
             records: list[list]) -> int:
    """Replace a lookup table with what the workbook carries."""
    cur = conn.cursor()
    cur.execute(f"DELETE FROM {table}")
    if not records:
        return 0
    cur.executemany(
        f"INSERT OR REPLACE INTO {table} ({', '.join(columns)}) "
        f"VALUES ({', '.join('?' for _ in columns)})", records)
    return len(records)


def import_status_map(conn, rows, header_row) -> int:
    h = rows[header_row - 1]
    a = column_index(h, "position status")
    b = column_index(h, "concur status")
    out = [[_cell(r[a]), _cell(r[b]) if b is not None and b < len(r) else ""]
           for r in rows[header_row:] if a is not None and a < len(r) and _cell(r[a])]
    return _refresh(conn, "ADP_Concur_StatusMap", ["position_status", "concur_status"], out)


def import_country_map(conn, rows, header_row) -> int:
    """
    The Country Map, which now answers three questions rather than one.

    ADP asks it only for the two-character code. UKG asks it for the currency
    as well, because UKG has no business unit and so cannot take the Org Map
    route ADP uses to reach one.
    """
    h = rows[header_row - 1]
    idx = {"adp_country": column_index(h, "legal / preferred address", "adp country"),
           "concur_country": column_index(h, "adp country", "concur country"),
           "country_name": column_index(h, "country"),
           "currency_code": column_index(h, "currency code"),
           "currency_name": column_index(h, "currency")}
    # 'ADP Country' is both the key on the old sheet and the value column on
    # the new one, so the two lookups above can land on the same column. When
    # they do, the value is the column after the key - which is what the
    # workbook's VLOOKUP(...,2) means.
    if (idx["adp_country"] is not None
            and idx["adp_country"] == idx["concur_country"]):
        idx["concur_country"] = idx["adp_country"] + 1
    key = idx["adp_country"]
    out = []
    for r in rows[header_row:]:
        if key is None or key >= len(r) or not _cell(r[key]):
            continue
        out.append([_cell(r[i]) if i is not None and i < len(r) else ""
                    for i in idx.values()])
    return _refresh(conn, "ADP_Concur_CountryMap", list(idx), out)


def import_country_ref(conn, rows, header_row) -> int:
    """
    The country reference block beside the Country Map.

    The old workbook kept this as a second block in columns E:H. The
    15 September sheet folded it into the map itself - one row per country
    carrying the code, the name, the currency code and the currency name - so
    this now reads the same five columns as import_country_map and keys them on
    the two-character Concur code rather than on ADP's three-character one.
    Same table, same contents; only the sheet changed shape underneath it.
    """
    h = rows[header_row - 1]
    key_at = column_index(h, "adp country")
    if key_at is None:
        return 0
    idx = {"country_code": key_at,
           "country_name": column_index(h, "country"),
           "currency_code": column_index(h, "currency code"),
           "currency_name": column_index(h, "currency")}
    out, seen = [], set()
    for r in rows[header_row:]:
        if key_at >= len(r):
            continue
        code = _cell(r[key_at]).strip().upper()
        if not code or code in seen:
            continue
        seen.add(code)
        row = [_cell(r[i]) if i is not None and i < len(r) else ""
               for i in idx.values()]
        row[0] = code
        out.append(row)
    return _refresh(conn, "ADP_Concur_CountryRef", list(idx), out)


def import_locale_map(conn, rows, header_row) -> int:
    """
    The locale block on the Language Map, columns M:O.

    Country code, the language's English name, and the locale Concur wants -
    'US' -> 'en_US'. It is a second table on the same sheet, keyed on something
    else entirely, which is why it is loaded as an extra block rather than as
    part of the Language Map.

    The sheet computes the country with =RIGHT(O2,2), the last two characters
    of the locale, so that is what is done here rather than trusting column M -
    M holds a formula and a workbook saved without recalculating would hand
    over a stale value or none at all.
    """
    out, seen = [], set()
    for r in rows[header_row:]:
        # M is index 12, N is 13, O is 14.
        if len(r) < 15:
            continue
        locale = _cell(r[14]).strip()
        name = _cell(r[13]).strip()
        if not locale or "_" not in locale:
            continue
        code = locale[-2:].upper()
        if code in seen:
            continue
        seen.add(code)
        out.append([code, name, locale])
    return _refresh(conn, "ADP_Concur_LocaleMap",
                    ["country_code", "locale_name", "locale_code"], out)


def import_role_map(conn, rows, header_row) -> int:
    """The Role Assignment Map - reference only; nothing derives from it yet."""
    h = rows[header_row - 1]
    idx = {"role": column_index(h, "role"),
           "category": column_index(h, "category"),
           "automatic": column_index(h, "assign role automatically")}
    key = idx["role"]
    out = []
    for r in rows[header_row:]:
        if key is None or key >= len(r) or not _cell(r[key]):
            continue
        out.append([_cell(r[i]) if i is not None and i < len(r) else ""
                    for i in idx.values()])
    return _refresh(conn, "ADP_Concur_RoleMap", list(idx), out)


def import_invoice_map(conn, rows, header_row) -> int:
    """
    The Invoice Exception Map: who gets Concur Invoice by name.

    The file numbers on this tab are typed by hand and some carry trailing
    spaces ('203011   '). Excel's VLOOKUP would miss those against a trimmed
    key, so they are trimmed on the way in and the map matches everybody it
    names - which is a small, deliberate difference from the spreadsheet.
    """
    h = rows[header_row - 1]
    idx = {"file_number": column_index(h, "invoice exception employee id",
                                       "invoice exception employee"),
           "employee_name": column_index(h, "invoice exception employee name"),
           "access": column_index(h, "invoice access value")}
    # Both the ID and the Name column start with 'Invoice Exception Employee',
    # so when the two resolve to the same place the name is the next column.
    if (idx["file_number"] is not None
            and idx["file_number"] == idx["employee_name"]):
        idx["employee_name"] = idx["file_number"] + 1
    key = idx["file_number"]
    out, seen = [], set()
    for r in rows[header_row:]:
        if key is None or key >= len(r):
            continue
        fn = _cell(r[key]).strip()
        if not fn or fn in seen:
            continue
        seen.add(fn)
        row = [_cell(r[i]) if i is not None and i < len(r) else ""
               for i in idx.values()]
        row[0] = fn
        row[2] = row[2].strip().upper()
        out.append(row)
    return _refresh(conn, "ADP_Concur_InvoiceMap", list(idx), out)


def import_org_map(conn, rows, header_row) -> int:
    h = rows[header_row - 1]
    idx = {
        "business_unit_desc": column_index(h, "business unit description"),
        "home_department_desc": column_index(h, "home department description"),
        "home_department_code": column_index(h, "home department code"),
        "org_unit_1": column_index(h, "concur org 1"),
        "org_unit_2": column_index(h, "concur / erp code org 2", "concur/erp code org 2"),
        "default_language": column_index(h, "default org language"),
        "currency": column_index(h, "reimbursement currency", "reumbursement currency"),
    }
    key = idx["business_unit_desc"]
    out = []
    for r in rows[header_row:]:
        if key is None or key >= len(r) or not _cell(r[key]):
            continue
        out.append([_cell(r[i]) if i is not None and i < len(r) else ""
                    for i in idx.values()])
    return _refresh(conn, "ADP_Concur_OrgMap", list(idx), out)


def import_language_map(conn, rows, header_row) -> int:
    """
    Only the first three columns are read.

    The tab carries a second, wider block off to the right - the full Concur
    locale list, 'English (Australia)' and friends - which is reference
    material for choosing the stem, not a lookup the workbook performs.
    """
    h = rows[header_row - 1]
    a = column_index(h, "language")
    b = column_index(h, "language code")
    c = column_index(h, "adp language")
    seen = set()
    out = []
    for r in rows[header_row:]:
        if a is None or a >= len(r):
            continue
        desc = _cell(r[a])
        if not desc or desc in seen:
            continue
        seen.add(desc)
        out.append([desc,
                    _cell(r[b]) if b is not None and b < len(r) else "",
                    _cell(r[c]) if c is not None and c < len(r) else ""])
    return _refresh(conn, "ADP_Concur_LanguageMap",
                    ["language_desc", "language_code", "adp_language"], out)


def import_salary_map(conn, rows, header_row) -> int:
    h = rows[header_row - 1]
    idx = {"pay_grade_code": column_index(h, "pay grade code"),
           "pay_grade_desc": column_index(h, "pay grade description"),
           "expense_map": column_index(h, "expense map"),
           "travel_map": column_index(h, "travel map"),
           "invoice_map": column_index(h, "invoice approval")}
    key = idx["pay_grade_code"]
    out = []
    for r in rows[header_row:]:
        if key is None or key >= len(r) or not _cell(r[key]):
            continue
        out.append([_cell(r[i]) if i is not None and i < len(r) else ""
                    for i in idx.values()])
    return _refresh(conn, "ADP_Concur_SalaryMap", list(idx), out)


def import_supervisor_map(conn, rows, header_row) -> int:
    """
    The supervisor exceptions.

    A duplicated employee row is not an error - the workbook has one - so the
    first entry for a file number wins and the rest are dropped, which is what
    VLOOKUP does. A blank supervisor is kept: it is how 'top of the food chain'
    is expressed, and the note column says so.
    """
    h = rows[header_row - 1]
    idx = {"file_number": column_index(h, "exception employee id"),
           "employee_name": column_index(h, "exception employee name"),
           "supervisor_name": column_index(h, "supervisor employee name"),
           "supervisor_id": column_index(h, "supervisor id")}
    note_at = 5  # the workbook keeps its free-text note in column F
    key = idx["file_number"]
    out, seen = [], set()
    for r in rows[header_row:]:
        if key is None or key >= len(r):
            continue
        fn = _cell(r[key])
        if not fn or fn in seen:
            continue
        seen.add(fn)
        row = [_cell(r[i]) if i is not None and i < len(r) else "" for i in idx.values()]
        row.append(_cell(r[note_at]) if note_at < len(r) else "")
        out.append(row)
    return _refresh(conn, "ADP_Concur_SupervisorMap",
                    list(idx) + ["note"], out)




def split_name(raw: str) -> tuple[str, str]:
    """
    'Shi, HanBing' -> ('Shi', 'HanBing').

    UKG sends one name field where ADP sends three, and the workbook splits it
    with =FIND(",",F2,1) / =LEFT(F2,BM2-1) / =MID(F2,BM2+2,45). Same split
    here, with one difference: a name with no comma keeps the whole string as
    the last name rather than becoming a #VALUE! error, because a person
    without a comma in their name is a person, not a broken row.
    """
    raw = _cell(raw).strip()
    if "," not in raw:
        return raw, ""
    last, _, first = raw.partition(",")
    return last.strip(), first.strip()


def import_ukg(conn: sqlite3.Connection, rows: list[tuple],
               header_row: int, import_id: int) -> dict:
    """
    Load the UKG export.

    Merged on File Number into the same table as ADP, because they are the
    same people in one company - the `source` column is what says which system
    sent a row, and it decides the derivation and the record types. There is
    no separate roster any more and no separate file.

    UKG has no duplicate-employment problem the way ADP does, so there is no
    row ranking here; the first row for a file number wins and any repeat is
    counted and reported. As with ADP, a field a hand edit changed away from
    what UKG sent - see _protected_set_clause() - survives this merge.
    """
    headings = list(rows[header_row - 1])
    positions = {}
    missing = []
    for heading, column in UKG_COLUMNS:
        idx = column_index(headings, heading)
        if idx is None:
            missing.append(heading)
        positions[column] = idx

    if positions.get("file_number") is None:
        raise ValueError("The UKG sheet has no 'Employee Number' column.")

    cols = [c for c in positions if positions[c] is not None]
    # The two names are derived from the one UKG sends, so they are written
    # alongside whatever the sheet gave us.
    write_cols = cols + ["legal_last_name", "legal_first_name"]
    insert_sql = (
        "INSERT INTO ADP_Concur_Employees (source, import_id, source_row, "
        + ", ".join(write_cols) + ") "
        "VALUES ('ukg', ?, ?, " + ", ".join("?" for _ in write_cols) + ") "
        "ON CONFLICT (file_number) DO UPDATE SET "
        + _protected_set_clause(write_cols)
        + ", source = 'ukg'"
        + ", import_id = excluded.import_id, source_row = excluded.source_row"
        + ", modified_at = datetime('now')"
        + ", row_state = CASE WHEN ADP_Concur_Employees.row_state = 'new' "
          "THEN 'new' ELSE 'unchanged' END"
    )

    cur = conn.cursor()
    loaded = skipped = duplicates = 0
    seen: set[str] = set()
    # Who was already here from ADP. Both sheets can name the same person - a
    # transfer between the two systems appears on both in the same cut - and
    # since both merge on File Number, the second load silently overwrites the
    # first. It is written onto the row as a note so it becomes an exception
    # rather than a surprise.
    from_adp = {r[0] for r in conn.execute(
        "SELECT file_number FROM ADP_Concur_Employees WHERE source = 'adp'")}
    overlap = []

    for n, r in enumerate(rows[header_row:], start=header_row + 1):
        fn_idx = positions["file_number"]
        file_number = _cell(r[fn_idx]) if fn_idx < len(r) else ""
        if not file_number:
            skipped += 1
            continue
        if file_number in seen:
            duplicates += 1
            continue
        seen.add(file_number)
        values = [_cell(r[positions[c]]) if positions[c] is not None
                  and positions[c] < len(r) else "" for c in cols]
        last, first = split_name(
            values[cols.index("employee_name_raw")]
            if "employee_name_raw" in cols else "")
        cur.execute(insert_sql, [import_id, n] + values + [last, first])
        if file_number in from_adp:
            overlap.append(file_number)
        loaded += 1

    for fn in overlap:
        cur.execute(
            "UPDATE ADP_Concur_Employees SET duplicate_note = ? "
            "WHERE file_number = ?",
            ("Both exports carry this File Number - the ADP row was replaced "
             "by the UKG one. Confirm which system owns this person.", fn))

    return {"loaded": loaded, "skipped": skipped, "duplicates": duplicates,
            "overlap": len(overlap), "missing_columns": missing}


def import_layout(conn: sqlite3.Connection, record_type: str,
                  rows: list[tuple], header_row: int) -> int:
    """
    Capture a record template's columns.

    The heading row is the one starting 'Trx Type (nnn)'. Where the template
    also carries a field-width row underneath - most tabs do - it is stored
    beside the heading, so the extract can be checked against the widths Concur
    publishes without going back to the spreadsheet.

    How wide the record is is decided by the *numbered* row above the headings,
    not by the last cell with text in it. Those numbers are the spec's field
    positions, and they are the only authority for where the record stops.
    The 700 tabs are why: both carry a note in the column after the last field
    - "Only load users who can work with invoices" - with no number and no
    width. Counting headings would make that note a seventeenth field, and
    since Concur requires every field to be represented, every 700 in the file
    would carry one delimiter too many.
    """
    headings = rows[header_row - 1]
    widths = rows[header_row] if len(rows) > header_row else []
    # A width row is all numbers-ish; a data row starts with the record type.
    if widths and _cell(widths[0]) == record_type:
        widths = []

    # The field-number row sits above the headings when the tab has one.
    numbers = rows[header_row - 2] if header_row >= 2 else []
    fields = 0
    for i, cell in enumerate(numbers):
        if _cell(cell).strip().isdigit() and int(_cell(cell)) == i + 1:
            fields = i + 1
        else:
            break

    cur = conn.cursor()
    cur.execute("DELETE FROM ADP_Concur_Layouts WHERE record_type = ?", (record_type,))
    out = []
    for i, heading in enumerate(headings):
        if fields and i >= fields:
            break
        if not fields and heading is None and i > 0 and all(
                h is None for h in headings[i:]):
            break
        out.append((record_type, i + 1, get_column_letter(i + 1),
                    " ".join(str(heading).split()) if heading is not None else "",
                    _cell(widths[i]) if i < len(widths) else ""))
    cur.executemany(
        "INSERT INTO ADP_Concur_Layouts "
        "(record_type, position, column_ref, heading, max_width) VALUES (?, ?, ?, ?, ?)",
        out)
    return len(out)


# ------------------------------------------------------------------- driver


# Blocks that live on the right-hand side of a tab whose left-hand side is
# already a map, keyed on something the left-hand block is not.
EXTRA_BLOCKS = {
    "country_map": [("country_ref", "ADP_Concur_CountryRef", None)],
    "language_map": [("locale_map", "ADP_Concur_LocaleMap", None)],
}

HANDLERS = {
    "status_map": ("ADP_Concur_StatusMap", import_status_map),
    "role_map": ("ADP_Concur_RoleMap", import_role_map),
    "invoice_map": ("ADP_Concur_InvoiceMap", import_invoice_map),
    "country_map": ("ADP_Concur_CountryMap", import_country_map),
    "org_map": ("ADP_Concur_OrgMap", import_org_map),
    "language_map": ("ADP_Concur_LanguageMap", import_language_map),
    "salary_map": ("ADP_Concur_SalaryMap", import_salary_map),
    "supervisor_map": ("ADP_Concur_SupervisorMap", import_supervisor_map),
}

EXTRA_BLOCKS["country_map"] = [("country_ref", "ADP_Concur_CountryRef",
                                import_country_ref)]
EXTRA_BLOCKS["language_map"] = [("locale_map", "ADP_Concur_LocaleMap",
                                 import_locale_map)]


def ADP_Concur_import_workbook(file_path: str | Path,
                               conn: sqlite3.Connection | None = None,
                               db_path: str | None = None,
                               derive: bool = True) -> dict:
    """
    Load one workbook. The whole thing runs in a single transaction.

    Returns what happened per worksheet, including the ones it did not
    recognise - a sheet that is not reported as loaded is one to look at.
    """
    path = Path(file_path)
    own_conn = conn is None
    conn = conn or connect(db_path)

    wb = load_workbook(path, read_only=True, data_only=True)
    cur = conn.cursor()
    cur.execute("INSERT INTO ADP_Concur_Imports (file_name, file_path) VALUES (?, ?)",
                (path.name, str(path)))
    import_id = cur.lastrowid

    sheets, unknown, ignored = [], [], []
    adp_sheets: list[tuple] = []
    employees = None
    ukg = None
    try:
        for ws in wb.worksheets:
            rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
            if not rows:
                continue
            kind, header_row = identify_sheet(rows, ws.title)
            if kind is None:
                unknown.append(ws.title)
                continue

            if kind == "employees":
                # Collected rather than loaded: a workbook can carry more than
                # one ADP-shaped sheet - this one has 'Multiple Payrolls', a
                # drill-down of the same people - and only the fullest cut is
                # the source. Loading a subset as well would re-run the
                # duplicate resolution over fewer rows and quietly change which
                # one won.
                adp_sheets.append((ws.title, rows, header_row))
                continue
                sheets.append({"sheet": ws.title, "kind": "ADP export",
                               "rows": employees["loaded"],
                               "source_rows": employees["source_rows"],
                               "duplicates": employees["duplicates"],
                               "skipped": employees["skipped"],
                               "missing_columns": employees["missing_columns"]})
            elif kind == "ukg":
                ukg = import_ukg(conn, rows, header_row, import_id)
                sheets.append({"sheet": ws.title, "kind": "UKG export",
                               "rows": ukg["loaded"],
                               "skipped": ukg["skipped"],
                               "duplicates": ukg["duplicates"],
                               "overlap": ukg["overlap"],
                               "missing_columns": ukg["missing_columns"]})
            elif kind.startswith("skip_"):
                ignored.append({
                    "sheet": ws.title,
                    "why": f"record type {kind.split('_', 1)[1]} is not in "
                           "scope for this load"})
            elif kind.startswith("record_"):
                record_type = kind.split("_", 1)[1]
                n = import_layout(conn, record_type, rows, header_row)
                sheets.append({"sheet": ws.title,
                               "kind": f"{record_type} layout", "rows": n})
            else:
                table, handler = HANDLERS[kind]
                n = handler(conn, rows, header_row)
                sheets.append({"sheet": ws.title, "kind": table, "rows": n})
                # Two of the tabs carry a second block keyed on something else
                # entirely, which the non-US formulas look up. Same sheet, same
                # header row, different table.
                for extra_kind, extra_table, extra_handler in EXTRA_BLOCKS.get(
                        kind, []):
                    extra_n = extra_handler(conn, rows, header_row)
                    if extra_n:
                        sheets.append({"sheet": ws.title, "kind": extra_table,
                                       "rows": extra_n})

        # The fullest ADP-shaped sheet is the source; the rest are working
        # copies. Counted on distinct File Numbers rather than rows, because a
        # drill-down of the duplicates can have more rows and fewer people.
        if adp_sheets:
            fn_at = lambda hs: column_index(list(hs), "File Number")
            def distinct(rows_, header_):
                i = fn_at(rows_[header_ - 1])
                if i is None:
                    return 0
                return len({_cell(r[i]) for r in rows_[header_:]
                            if i < len(r) and _cell(r[i])})
            # Excel writes "Details for ..." into A1 of a sheet it generates by
            # double-clicking a pivot total, which is exactly what 'Multiple
            # Payrolls' is. That marker is what separates a generated copy from
            # the real export when the two hold the same people - and here they
            # hold all 166 either way, so counting alone would pick by luck.
            drilldown = lambda rows_: _norm(
                rows_[0][0] if rows_ and rows_[0] else "").startswith("details for")
            ranked = sorted(adp_sheets,
                            key=lambda t: (drilldown(t[1]), -distinct(t[1], t[2])))
            title, rows_, header_ = ranked[0]
            employees = import_employees(conn, rows_, header_, import_id)
            sheets.insert(0, {"sheet": title, "kind": "ADP export",
                              "rows": employees["loaded"],
                              "source_rows": employees["source_rows"],
                              "duplicates": employees["duplicates"],
                              "skipped": employees["skipped"],
                              "missing_columns": employees["missing_columns"]})
            for other_title, other_rows, other_header in ranked[1:]:
                ignored.append({
                    "sheet": other_title,
                    "why": f"a second ADP-shaped sheet with "
                           f"{distinct(other_rows, other_header)} people against "
                           f"{distinct(rows_, header_)} on '{title}'"
                           + (" - a pivot drill-down, not the export"
                              if drilldown(other_rows)
                              else " - the fullest cut is the source")})

        cur.execute(
            "UPDATE ADP_Concur_Imports SET sheet_counts = ?, row_count = ? "
            "WHERE import_id = ?",
            (json.dumps(sheets), employees["loaded"] if employees else 0, import_id))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        wb.close()

    result = {"file": path.name, "import_id": import_id, "sheets": sheets,
              "unknown_sheets": unknown, "ignored_sheets": ignored,
              "employees": employees["loaded"] if employees else 0,
              "ukg": ukg["loaded"] if ukg else 0}

    if derive:
        from ADP_Concur_Map import ADP_Concur_derive, load_config
        result["derive"] = ADP_Concur_derive(conn, load_config())

    if own_conn:
        conn.close()
    return result


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Load an ADP/Concur workbook.")
    ap.add_argument("files", nargs="+", help="One or more .xlsx workbooks")
    ap.add_argument("--db", default=None, help=f"SQLite path (default {DEFAULT_DB_PATH})")
    ap.add_argument("--no-derive", action="store_true",
                    help="Skip rebuilding the derived columns afterwards")
    args = ap.parse_args(argv)

    conn = connect(args.db)
    print(f"Database: {resolve_db_path(args.db)}")
    for f in args.files:
        res = ADP_Concur_import_workbook(f, conn=conn, derive=not args.no_derive)
        print(f"\n{res['file']}  (import {res['import_id']})")
        for s in res["sheets"]:
            extra = ""
            if s.get("duplicates"):
                extra += (f" from {s['source_rows']} ADP row(s), "
                          f"{s['duplicates']} employee(s) had more than one")
            if s.get("skipped"):
                extra += f", {s['skipped']} unkeyed row(s) skipped"
            if s.get("missing_columns"):
                extra += f", missing: {', '.join(s['missing_columns'])}"
            print(f"  {s['sheet']:<18} {s['kind']:<26} {s['rows']:>5} row(s){extra}")
        for u in res["unknown_sheets"]:
            print(f"  {u:<18} not recognised - ignored")
        for i in res["ignored_sheets"]:
            print(f"  {i['sheet']:<18} ignored - {i['why']}")
        if "derive" in res:
            d = res["derive"]
            print(f"  derived {d['employees']} employee(s): "
                  f"{d['errors']} error(s), {d['warnings']} warning(s)")
    conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
