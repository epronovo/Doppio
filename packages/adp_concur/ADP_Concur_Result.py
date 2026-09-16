#!/usr/bin/env python3
"""
ADP_Concur_Result - read Concur's load result back in and put names to it.

Concur answers a load with a spreadsheet of three columns: Level, Record
Identifier, Message. The Record Identifier is a *line number in the file you
sent*, which makes the result unreadable on its own - 'Record 139' is nobody.
This module joins that line number back through the extract to the record type
and the employee, classifies each message by cause, and stores the lot, so the
next run starts from 'these eleven people' rather than from a spreadsheet.

The join is the whole point. Everything else here is bookkeeping.

Run it from the command line:

    ADP_Concur_Result.py --load 'Employee_p0010945e24e Run-18.xls'
    ADP_Concur_Result.py --load result.xls --extract 12   # pick the extract
    ADP_Concur_Result.py --show                           # the last load
    ADP_Concur_Result.py --runs                           # every load
    ADP_Concur_Result.py --clear CLEAR                    # drop them all
"""

from __future__ import annotations

import argparse
import os
import re
import sqlite3
import sys
from pathlib import Path

from ADP_Concur_Db import connect, resolve_db_path

# --------------------------------------------------------------- reading .xls
#
# Concur sends BIFF (.xls), which openpyxl cannot open, so xlrd does the old
# format and openpyxl the new one. xlrd 2.x dropped .xlsx deliberately - it
# only reads .xls now - which is why both are here rather than one.


XLS_HELP = (
    "Reading Concur's .xls needs the 'xlrd' package, which is not installed:\n"
    "    pip install xlrd\n"
    "Or, without installing anything: open the file in Excel and Save As "
    "'.xlsx', then drop that instead - the app reads both.")


def xls_supported() -> bool:
    """Whether this machine can read the old .xls format at all."""
    try:
        import xlrd                                    # noqa: F401, PLC0415
    except ImportError:
        return False
    return True


def _rows_from_xls(path: Path) -> list[list[str]]:
    # Concur only ever sends BIFF, so this is the common path, not the exotic
    # one - and 'No module named xlrd' is a useless thing to show somebody who
    # just dropped a file on a web page. The error says what to install and
    # what to do instead of installing it.
    try:
        import xlrd  # noqa: PLC0415 - optional, and only for the old format
    except ImportError as exc:
        raise RuntimeError(f"{path.name}: {XLS_HELP}") from exc
    sheet = xlrd.open_workbook(str(path)).sheet_by_index(0)
    return [[str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
            for r in range(sheet.nrows)]


def _rows_from_xlsx(path: Path) -> list[list[str]]:
    import openpyxl  # noqa: PLC0415
    ws = openpyxl.load_workbook(str(path), data_only=True, read_only=True).worksheets[0]
    return [["" if c is None else str(c).strip() for c in row]
            for row in ws.iter_rows(values_only=True)]


def read_result_rows(path: Path | str) -> list[tuple[str, str, str]]:
    """
    The (level, record identifier, message) triples out of a result workbook.

    The header row is found by content rather than assumed to be row 1, for the
    same reason the workbook importer does it: a file that grows a title row
    should not silently lose its first record.
    """
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".xls":
        rows = _rows_from_xls(path)
    elif suffix in (".xlsx", ".xlsm"):
        rows = _rows_from_xlsx(path)
    else:
        raise ValueError(f"{path.name}: expected a .xls or .xlsx result file.")

    start = 0
    for i, row in enumerate(rows[:20]):
        joined = " ".join(str(c).lower() for c in row)
        if "record identifier" in joined and "message" in joined:
            start = i + 1
            break

    out = []
    for row in rows[start:]:
        cells = list(row) + ["", "", ""]
        level, record_id, message = (str(cells[0]).strip(),
                                     str(cells[1]).strip(),
                                     str(cells[2]).strip())
        if not message:
            continue
        # xlrd hands back floats for numeric cells: '139.0' is line 139.
        record_id = re.sub(r"\.0$", "", record_id)
        out.append((level or "Info", record_id, message))
    return out


def looks_like_a_result(path: Path | str) -> bool:
    """
    Whether this file is Concur talking back, rather than the ADP workbook.

    Decided by content, like everything else that identifies a sheet here: a
    .xls can only be a result, and a .xlsx is one if its first sheet is headed
    the way Concur heads them. Names are no good - Concur's are machine
    generated ('Employee_p0010945e24e Run-18') and nobody keeps them.
    """
    path = Path(path)
    if path.suffix.lower() == ".xls":
        return True
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return False
    try:
        for row in _rows_from_xlsx(path)[:20]:
            joined = " ".join(str(c).lower() for c in row)
            if "record identifier" in joined and "message" in joined:
                return True
    except Exception:                                  # noqa: BLE001
        return False
    return False


# --------------------------------------------------------- classifying causes
#
# Each entry is (category, severity hint, matcher, what to do about it). The
# category is what the Fixes view groups on, so the wording is the wording a
# person reads. Order matters - the first match wins - so the specific
# patterns sit above the general ones.

CATEGORIES: list[tuple[str, str, re.Pattern, str]] = [
    ("Missing country code", "error",
     re.compile(r"missing required field\s*-\s*ctry_code", re.I),
     "The 305 was rejected outright. Concur requires Ctry Code when it is "
     "creating the employee, so this hits new starters and nobody else. "
     "Column J is filled from the derived country - if this comes back, the "
     "person has no country to derive one from."),

    ("Travel rule class not recognised", "error",
     re.compile(r"rule class '([^']*)' is invalid", re.I),
     "The 350's Travel Class Name is not a rule class in this Concur tenant. "
     "It is fed from the Salary Map's Travel column. Either the tenant needs "
     "those rule classes, or the 350 should not be in the load at all - the "
     "July file that loaded cleanly carried none."),

    ("Travel record with no employee behind it", "error",
     re.compile(r"travel information .* cannot be imported", re.I),
     "The 350 landed but its 305 did not, so there was nothing to attach it "
     "to. Fix the 305 error for the same person and this goes with it."),

    ("Business unit does not resolve", "error",
     re.compile(r"employeehierarchyservice|hierarchy node", re.I),
     "Concur could not place the person in the expense hierarchy. It is the "
     "Org Map: the segment value in the message is what we sent. Map the BU, "
     "or drop the person from the load if they have left."),

    ("Org Unit 2 is not a list code", "error",
     re.compile(r"connected list.*invalid list code|invalid list code.*OrgUnit",
                re.I),
     "Concur has Org Unit 2 on the connected list '*Division - Department' and "
     "the value sent is not in it, so the record was refused. The US roster "
     "derives a numeric code from the Org Map and passes; the non-US tabs have "
     "a department name typed in ('Production', 'SALES', 'G&A') and will not. "
     "Those tabs need the real codes - blank is accepted, a name is not."),

    ("Value is not in a Concur list", "warning",
     re.compile(r"could not be resolved to an existing custom list item", re.I),
     "The record loaded but the field was dropped. Either the connected list "
     "in Concur is missing these items, or the column should be blank - set "
     "it in blank_fields on the Extract tab to stop sending it."),

    ("Login ID differs from the one on file", "warning",
     re.compile(r"login id provided .* is different than currently exists", re.I),
     "The person already exists in Concur under another Login ID. A 305 "
     "cannot change it - that needs a 320 record or User Administration. "
     "Expected while the '.new.uat' suffix is in play."),

    ("Approver is not in Concur", "warning",
     re.compile(r"approver could not be assigned", re.I),
     "The approver is not loaded yet. Usually the other roster: an approver "
     "on the non-US file cannot be resolved by the US file that loads first. "
     "Load the other roster and run again."),

    ("Manager loop inside Concur", "warning",
     re.compile(r"circular reference", re.I),
     "Concur's own data has a loop. We write BI Manager blank and the 100 "
     "record says UPDATE, which never clears a field - so a stale manager "
     "survives every load. Populate BI Manager, or fix the profile by hand."),

    ("Manager is not a valid employee", "warning",
     re.compile(r"bi manager .* is not a valid employee", re.I),
     "The manager is not in Concur. Same cause as the approver warnings - "
     "usually somebody on the other roster."),
]

FALLBACK = ("Other", "", "Not one of the causes this app knows about yet. "
                         "Worth reading in full.")


def classify(message: str) -> tuple[str, str]:
    """(category, what to do) for one Concur message."""
    for name, _sev, pattern, remedy in CATEGORIES:
        if pattern.search(message):
            return name, remedy
    return FALLBACK[0], FALLBACK[2]


FIELD_RE = re.compile(r"Field:\s*([A-Za-z0-9_ ]+?)\s+Value:", re.I)


def field_in(message: str) -> str:
    """The Concur field name when the message names one."""
    m = FIELD_RE.search(message)
    return m.group(1).strip() if m else ""


# ------------------------------------------------------ the extract line index


def index_extract(path: Path | str) -> dict[int, tuple[str, str]]:
    """
    line number -> (record type, employee id), for one extract file.

    Concur counts every line including the 100 record, so line 1 is the import
    settings and the employee records start at 2. The employee id is field 5 on
    a 305 and field 2 on everything else, which is the only per-record-type
    knowledge this needs.
    """
    import csv  # noqa: PLC0415

    out: dict[int, tuple[str, str]] = {}
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for n, row in enumerate(csv.reader(fh), start=1):
            if not row:
                continue
            record_type = row[0].strip()
            if record_type == "305":
                emp = row[4].strip() if len(row) > 4 else ""
            elif record_type == "100":
                emp = ""
            else:
                emp = row[1].strip() if len(row) > 1 else ""
            out[n] = (record_type, emp)
    return out


EMPLOYEE_ID_RE = re.compile(
    r"(?:Employee_ID|EmployeeId|employee ID|Employee ID)[:=\s]*\[?\s*'?([A-Za-z0-9]+)'?",
    re.I)


def quoted_employee_id(message: str) -> str:
    """
    The employee Concur names in a message, when it names one.

    Two things make this fussier than it looks. The token must contain a digit,
    because "employee ID for approver does not exist" matches the prefix and
    would otherwise hand back the word 'for'. And it is the *last* match that
    counts, not the first: where a message names two people it names the
    subject second - "The BI Manager ID '207199' cannot be assigned to employee
    ID '900848'" is about 900848, and taking the first match would file it
    against the manager instead.
    """
    found = [m.group(1).strip() for m in EMPLOYEE_ID_RE.finditer(message)]
    for token in reversed(found):
        if any(ch.isdigit() for ch in token):
            return token
    return ""


def score_extract(lines: dict[int, tuple[str, str]],
                  rows: list[tuple[str, str, str]]) -> tuple[int, int]:
    """
    How well one extract explains a result file: (agreements, disagreements).

    Line count alone is a weak match - re-export the same people and you get
    another file of exactly the same size - so the join is checked rather than
    assumed. Concur quotes the employee in many of its messages, and that
    employee has to be the one sitting on the line it is complaining about.
    Every message that carries an ID is a free assertion about the join, and a
    single disagreement means the file being read is not the file that was
    sent.
    """
    agree = disagree = 0
    for _level, record_id, message in rows:
        quoted = quoted_employee_id(message)
        if not (quoted and record_id.isdigit()):
            continue
        found = lines.get(int(record_id), ("", ""))[1]
        if not found:
            continue
        if found == quoted or found.lstrip("0") == quoted.lstrip("0"):
            agree += 1
        else:
            disagree += 1
    return agree, disagree


def join_is_sound(agree: int, disagree: int) -> bool:
    """
    Whether the agreements are convincing enough to trust the names.

    Not zero-tolerance, because Concur itself is not exact: in the Run-18
    result four messages carried line 369 - the last line of the file - while
    naming two people who sit hundreds of lines earlier. A handful of those is
    Concur mis-stamping, and the messages still name the right person. A wrong
    extract looks nothing like that: it disagrees everywhere at once. One in
    twenty is comfortably between the two.
    """
    total = agree + disagree
    return total > 0 and disagree <= max(2, total * 0.05)


def find_extract(conn: sqlite3.Connection, highest_line: int,
                 rows: list[tuple[str, str, str]],
                 extract_key: int | None = None) -> tuple[int, str, str, dict]:
    """
    Which extract a result file is answering: (key, name, path, how it matched).

    Candidates are the extracts of the right length whose file is still on
    disk; the winner is the one whose employee IDs agree with Concur's and
    disagree with it least. An explicit --extract skips the contest but is
    still scored, so choosing the wrong one by hand says so rather than
    quietly mislabelling every error in the run.
    """
    sql = ("SELECT extract_key, file_name, file_path, "
           "       1 + n_305 + n_350 + n_360 AS lines "
           "FROM ADP_Concur_Extracts "
           + ("WHERE extract_key = ? " if extract_key else "")
           + "ORDER BY written_at DESC, extract_key DESC")
    candidates = list(conn.execute(sql, (extract_key,) if extract_key else ()))

    best = None
    for row in candidates:
        key, name, path, n_lines = row[0], row[1], row[2] or "", row[3]
        if not path or not os.path.exists(path):
            continue
        if not extract_key and n_lines != highest_line:
            continue
        agree, disagree = score_extract(index_extract(path), rows)
        cand = {"key": key, "name": name, "path": path,
                "agree": agree, "disagree": disagree, "lines": n_lines}
        if best is None or (agree - 5 * disagree) > (best["agree"] - 5 * best["disagree"]):
            best = cand

    if not best:
        return 0, "", "", {"agree": 0, "disagree": 0, "candidates": len(candidates)}
    return best["key"], best["name"], best["path"], best


# ------------------------------------------------------------------- the load


def ADP_Concur_load_result(conn: sqlite3.Connection, path: Path | str,
                           run_label: str = "",
                           extract_key: int | None = None) -> dict:
    """
    Read one Concur result file into ADP_Concur_Loads / ADP_Concur_Results.

    Everything happens in one transaction: a half-read result is worse than
    none, because the counts would lie.
    """
    path = Path(path)
    rows = read_result_rows(path)
    if not rows:
        raise ValueError(f"{path.name}: no Level / Record Identifier / Message "
                         f"rows found.")

    numbers = [int(r[1]) for r in rows if r[1].isdigit()]
    highest = max(numbers) if numbers else 0
    key, extract_name, extract_path, match = find_extract(
        conn, highest, rows, extract_key)
    lines = index_extract(extract_path) if extract_path else {}

    # file number -> (employee_key, name), so a result row can name a person
    # even when they have since been edited.
    people = {
        str(r[0]): (r[1], f"{r[2]}, {r[3]}".strip(", "))
        for r in conn.execute(
            "SELECT file_number, employee_key, legal_last_name, legal_first_name "
            "FROM ADP_Concur_Employees WHERE row_state <> 'deleted'")}

    label = run_label or re.sub(r"\.(xls|xlsx|xlsm)$", "", path.name, flags=re.I)

    with conn:
        cur = conn.execute(
            "INSERT INTO ADP_Concur_Loads "
            "(run_label, source_file, extract_key, extract_name, "
            " n_error, n_warning, n_message) VALUES (?,?,?,?,0,0,0)",
            (label, str(path), key or None, extract_name))
        load_key = cur.lastrowid

        n_error = n_warning = 0
        batch = []
        for level, record_id, message in rows:
            line = int(record_id) if record_id.isdigit() else 0
            record_type, at_line = lines.get(line, ("", ""))
            # Concur's own words beat its line number. The Record Identifier is
            # usually right, but not always: in the Run-18 result four messages
            # about 900848 and 211181 were all stamped with line 369, the last
            # line of the file, which belongs to somebody else entirely. Where
            # the message says who it is about, that is who it is about; the
            # line still supplies the record type.
            file_number = quoted_employee_id(message) or at_line
            employee_key, employee_name = people.get(file_number, (None, ""))
            category, _remedy = classify(message)
            if level.lower().startswith("err"):
                n_error += 1
            elif level.lower().startswith("warn"):
                n_warning += 1
            batch.append((load_key, level, line, record_type, file_number,
                          employee_key, employee_name, category,
                          field_in(message), message))

        conn.executemany(
            "INSERT INTO ADP_Concur_Results "
            "(load_key, level, record_id, record_type, file_number, "
            " employee_key, employee_name, category, field, message) "
            "VALUES (?,?,?,?,?,?,?,?,?,?)", batch)
        conn.execute(
            "UPDATE ADP_Concur_Loads SET n_error = ?, n_warning = ?, "
            "n_message = ? WHERE load_key = ?",
            (n_error, n_warning, len(batch), load_key))

    return {
        "load_key": load_key,
        "run_label": label,
        "messages": len(rows),
        "errors": n_error,
        "warnings": n_warning,
        "extract_key": key,
        "extract_name": extract_name,
        "matched": bool(lines),
        "highest_line": highest,
        "people": len({b[4] for b in batch if b[4]}),
        "agree": match.get("agree", 0),
        "disagree": match.get("disagree", 0),
        "join_ok": join_is_sound(match.get("agree", 0), match.get("disagree", 0)),
    }


# --------------------------------------------------------------- reading back


def ADP_Concur_result_summary(conn: sqlite3.Connection,
                              load_key: int | None = None) -> dict:
    """One load, grouped the way somebody would work through it."""
    row = conn.execute(
        "SELECT * FROM ADP_Concur_Loads "
        + ("WHERE load_key = ? " if load_key else "")
        + "ORDER BY loaded_at DESC, load_key DESC LIMIT 1",
        (load_key,) if load_key else ()).fetchone()
    if not row:
        return {"load": None, "groups": [], "unmatched": 0}

    load = dict(row)
    key = load["load_key"]
    remedies = {name: remedy for name, _s, _p, remedy in CATEGORIES}
    remedies[FALLBACK[0]] = FALLBACK[2]

    groups = []
    for g in conn.execute(
            "SELECT category, level, COUNT(*) AS n, "
            "       COUNT(DISTINCT file_number) AS people, "
            "       GROUP_CONCAT(DISTINCT record_type) AS record_types "
            "FROM ADP_Concur_Results WHERE load_key = ? "
            "GROUP BY category, level ORDER BY (level = 'Error') DESC, n DESC",
            (key,)):
        g = dict(g)
        g["remedy"] = remedies.get(g["category"], "")
        g["rows"] = [dict(r) for r in conn.execute(
            "SELECT record_id, record_type, file_number, employee_name, "
            "       field, message FROM ADP_Concur_Results "
            "WHERE load_key = ? AND category = ? AND level = ? "
            "ORDER BY record_id", (key, g["category"], g["level"]))]
        groups.append(g)

    unmatched = conn.execute(
        "SELECT COUNT(*) FROM ADP_Concur_Results "
        "WHERE load_key = ? AND file_number = ''", (key,)).fetchone()[0]
    return {"load": load, "groups": groups, "unmatched": unmatched}


def ADP_Concur_result_people(conn: sqlite3.Connection,
                             load_key: int | None = None) -> list[dict]:
    """The same load seen per person - who to go and fix."""
    if load_key is None:
        row = conn.execute("SELECT load_key FROM ADP_Concur_Loads "
                           "ORDER BY loaded_at DESC, load_key DESC "
                           "LIMIT 1").fetchone()
        if not row:
            return []
        load_key = row[0]
    return [dict(r) for r in conn.execute(
        "SELECT file_number, "
        "       MAX(employee_name) AS employee_name, "
        "       MAX(employee_key)  AS employee_key, "
        "       SUM(level = 'Error')   AS errors, "
        "       SUM(level = 'Warning') AS warnings, "
        "       GROUP_CONCAT(DISTINCT category) AS categories "
        "FROM ADP_Concur_Results "
        "WHERE load_key = ? AND file_number <> '' "
        "GROUP BY file_number "
        "ORDER BY errors DESC, warnings DESC, employee_name", (load_key,))]


def ADP_Concur_runs(conn: sqlite3.Connection) -> list[dict]:
    """Every result file loaded, newest first."""
    return [dict(r) for r in conn.execute(
        "SELECT * FROM ADP_Concur_Loads "
        "ORDER BY loaded_at DESC, load_key DESC")]


def ADP_Concur_clear_results(conn: sqlite3.Connection,
                             load_key: int | None = None) -> dict:
    """Drop one load, or all of them."""
    with conn:
        if load_key:
            n = conn.execute("SELECT COUNT(*) FROM ADP_Concur_Results "
                             "WHERE load_key = ?", (load_key,)).fetchone()[0]
            conn.execute("DELETE FROM ADP_Concur_Results WHERE load_key = ?",
                         (load_key,))
            conn.execute("DELETE FROM ADP_Concur_Loads WHERE load_key = ?",
                         (load_key,))
        else:
            n = conn.execute("SELECT COUNT(*) FROM ADP_Concur_Results").fetchone()[0]
            conn.execute("DELETE FROM ADP_Concur_Results")
            conn.execute("DELETE FROM ADP_Concur_Loads")
    return {"deleted": n}


# -------------------------------------------------------------------- the CLI


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--db", default=None,
                    help="SQLite path (default from ADP_CONCUR_DB)")
    ap.add_argument("--load", metavar="FILE",
                    help="a Concur result workbook (.xls or .xlsx)")
    ap.add_argument("--label", default="", help="name this run")
    ap.add_argument("--extract", type=int, default=None,
                    help="extract_key to join against, when the size match is wrong")
    ap.add_argument("--show", action="store_true", help="the last load, by cause")
    ap.add_argument("--people", action="store_true", help="the last load, by person")
    ap.add_argument("--runs", action="store_true", help="every load")
    ap.add_argument("--clear", metavar="CLEAR",
                    help="type CLEAR to delete every stored result")
    args = ap.parse_args(argv)

    conn = connect(args.db)
    print(f"Database: {resolve_db_path(args.db)}")

    if args.clear:
        if args.clear != "CLEAR":
            print("Refusing: pass --clear CLEAR to confirm.", file=sys.stderr)
            return 2
        print(f"Deleted {ADP_Concur_clear_results(conn)['deleted']} result row(s).")
        return 0

    if args.load:
        r = ADP_Concur_load_result(conn, args.load, args.label, args.extract)
        print(f"Loaded {r['run_label']}: {r['messages']} message(s) - "
              f"{r['errors']} error(s), {r['warnings']} warning(s).")
        if r["matched"]:
            print(f"Joined to extract #{r['extract_key']} {r['extract_name']} - "
                  f"{r['people']} people named.")
            print(f"  Join checked against the employee IDs Concur quoted: "
                  f"{r['agree']} agree, {r['disagree']} disagree.")
            if not r["join_ok"]:
                print("  That is too many to be Concur mis-stamping a line. "
                      "This result probably answers a different file - re-run "
                      "with --extract to pick the right one.", file=sys.stderr)
        elif r["people"]:
            # Not a failure. Concur names the employee in most of its messages,
            # so the people land even with no extract to join to - what is lost
            # is the record type, and the few messages that name nobody.
            print(f"  No extract of {r['highest_line']} lines to join to, so "
                  f"the record types are blank. {r['people']} people were still "
                  f"named, from the employee IDs Concur quotes.")
        else:
            print(f"  No extract of {r['highest_line']} lines to join to, and "
                  f"no employee IDs in the messages - these are record numbers "
                  f"only. Pass --extract to choose one.")

    if args.runs:
        for run in ADP_Concur_runs(conn):
            print(f"  #{run['load_key']:<4} {run['loaded_at']}  "
                  f"{run['run_label'][:44]:46s} "
                  f"{run['n_error']:>4} err  {run['n_warning']:>4} warn")

    if args.people:
        for p in ADP_Concur_result_people(conn):
            print(f"  {p['file_number']:<10} {(p['employee_name'] or '?')[:30]:32s}"
                  f" {p['errors']:>3} err {p['warnings']:>3} warn  "
                  f"{p['categories']}")

    if args.show or not (args.load or args.runs or args.people):
        s = ADP_Concur_result_summary(conn)
        if not s["load"]:
            print("No results loaded yet.")
            return 0
        load = s["load"]
        print(f"\n{load['run_label']} - {load['n_error']} error(s), "
              f"{load['n_warning']} warning(s)\n")
        for g in s["groups"]:
            print(f"  [{g['level']:<7}] {g['category']:<40} "
                  f"{g['n']:>4} message(s), {g['people']:>3} people "
                  f"({g['record_types'] or '-'})")
        if s["unmatched"]:
            print(f"\n  {s['unmatched']} message(s) could not be tied to a person.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
