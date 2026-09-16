"""
ERP_Concur_Results - read the import results report Concur sends back, and
work out which purchase orders have to go again.

The report is the other half of this tool. The rest of the package cuts a file
down and sends it; this module reads the answer, and the answer is what makes
the next cut: drop the report on the page and every order Concur rejected is
picked, ready to be written out as a file of its own.

Three columns come back - Level, Record Identifier, Message - and neither of
the useful two says which purchase order it is about.

**The Record Identifier is a line number.** It is the 1-based position of the
record in the file that was sent, and the Info row that closes the report
carries the line count of the whole file. So a report can be resolved against
the purchase order file it came from and nothing else: line 272 of
purchase_order_import_<entity>_<stamp>.txt is the 200 header of the order the
row at identifier 272 is complaining about. Records that are not headers
resolve through the same positional grouping the parser uses everywhere else -
the 210 at line 3235 belongs to the header at 3238, so both rows of a
"sequence of the record types is invalid" pair land on one order.

**Only some messages name the order.** Error 5001 and error 4002 carry a
purchase order number in their text; the sequence pair carries nothing at all.
Both readings are therefore kept side by side - `stated_po` from the text,
`resolved_po` from the line number - and where both are known they are
compared. That comparison is the whole safety net: a report read against the
wrong run still resolves, to real orders that happen to sit at those lines,
and nothing else in the file would say so. `aligned` says so.

Nothing here writes a file or changes a record. It picks, and picking is the
selection the extract already reads.
"""
from __future__ import annotations

import csv
import hashlib
import html
import io
import json
import logging
import re
import sqlite3
from pathlib import Path

from ERP_Concur_Db import RESULTS_KIND, RESULTS_LABEL, replace_file
from ERP_Concur_Parse import ParseError

log = logging.getLogger("ERP_Concur_Results")

# What the report is saved as. Concur's own export is .xls; people re-save it
# as .xlsx, and the same three columns turn up pasted into a csv often enough
# to be worth reading.
SPREADSHEET_SUFFIXES = {".xls", ".xlsx", ".xlsm"}
TEXT_SUFFIXES = {".csv", ".tsv", ".txt"}

# The column headings, lower-cased. The report has been seen with and without
# a title row above them, so the header row is found by looking for this rather
# than by taking row 1.
COL_LEVEL = "level"
COL_RECORD = "record identifier"
COL_MESSAGE = "message"


class ResultsError(ParseError):
    """A file that cannot be read as a Concur import results report."""


# ----------------------------------------------------------------- reading


def _as_int(value) -> int | None:
    """
    The Record Identifier as a number, whatever the sheet stored it as.

    xlrd hands back 272.0 for a cell typed as a number and '272' for one typed
    as text, and a report pasted into a csv gives '272 '. All three are the
    same line.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def _grid_xls(path: Path) -> list[list]:
    try:
        import xlrd
    except ImportError as exc:                                   # noqa: BLE001
        raise ResultsError(
            "Reading a .xls results report needs xlrd - pip install xlrd. "
            "Saving the report as .xlsx or .csv works without it.") from exc
    book = xlrd.open_workbook(str(path))
    sheet = book.sheet_by_index(0)
    return [[sheet.cell_value(r, c) for c in range(sheet.ncols)]
            for r in range(sheet.nrows)]


def _grid_xlsx(path: Path) -> list[list]:
    from openpyxl import load_workbook
    book = load_workbook(str(path), read_only=True, data_only=True)
    sheet = book[book.sheetnames[0]]
    return [list(row) for row in sheet.iter_rows(values_only=True)]


def _grid_text(path: Path) -> list[list]:
    text = path.read_bytes().decode("utf-8-sig", errors="replace")
    head = text.split("\n", 1)[0]
    delimiter = "\t" if head.count("\t") > head.count(",") else ","
    return [row for row in csv.reader(io.StringIO(text), delimiter=delimiter)]


def read_grid(path: Path) -> list[list]:
    """
    The report as a grid of cells, whatever it was saved as.

    The extension decides which reader is tried first, but not which one is
    used: a .xls that is really an HTML table (Concur has been known to mail
    one) fails the xlrd read, and the text reader gets a turn before the file
    is refused.
    """
    path = Path(path)
    suffix = path.suffix.lower()
    order = ([_grid_xls, _grid_xlsx, _grid_text] if suffix == ".xls"
             else [_grid_xlsx, _grid_xls, _grid_text] if suffix in SPREADSHEET_SUFFIXES
             else [_grid_text, _grid_xlsx, _grid_xls])
    last = None
    for reader in order:
        try:
            grid = reader(path)
        except ResultsError:
            raise
        except Exception as exc:                                 # noqa: BLE001
            last = exc
            continue
        if grid:
            return grid
    raise ResultsError(f"{path.name} could not be read as a spreadsheet or a "
                       f"delimited file ({last}).")


def find_header(grid: list[list]) -> tuple[int, dict[str, int]]:
    """
    Which row is the heading row, and which column is which.

    Matched on the heading text rather than on position, because the report
    arrives with a title row above the headings as often as not, and because
    the column order is not worth trusting to stay put.
    """
    for i, row in enumerate(grid[:25]):
        cells = {str(c).strip().lower(): j for j, c in enumerate(row)
                 if c is not None and str(c).strip()}
        if COL_RECORD in cells and COL_MESSAGE in cells:
            return i, {"level": cells.get(COL_LEVEL, -1),
                       "record": cells[COL_RECORD],
                       "message": cells[COL_MESSAGE]}
    seen = ", ".join(sorted({str(c).strip() for row in grid[:5] for c in row
                             if c is not None and str(c).strip()})[:8]) or "nothing"
    raise ResultsError(
        "This does not look like a Concur import results report - no row of "
        f"headings with 'Record Identifier' and 'Message' in it. Found: {seen}.")


# ----------------------------------------------------------------- messages

MARKUP_RE = re.compile(r"<[^>]{0,40}>")
SPACE_RE = re.compile(r"\s+")

# The order number has to start with a digit. Without that, "the sequence of
# the record types for a purchase order is invalid" reads as an order called
# "is", and "a purchase order record" as one called "record" - both of which
# resolve to nothing and quietly mask the two messages that matter most,
# because those are exactly the ones that name no order at all.
PO_RE = re.compile(r"purchase order\s+(\d[A-Za-z0-9._/\-]*?)[.\s,;:]",
                   re.IGNORECASE)
CODE_RE = re.compile(r"Error Code:\s*([A-Za-z0-9\-]+)", re.IGNORECASE)
ERRTEXT_RE = re.compile(
    r"Error Message:\s*(.*?)(?=\s*(?:Field Level:|Field Code:|"
    r"Line Item External Id:|Error Code:)|\s*$)", re.IGNORECASE | re.DOTALL)
FIELD_LEVEL_RE = re.compile(r"Field Level:\s*([^;]+?)(?=\s*(?:;|Field Code:|$))",
                            re.IGNORECASE)
FIELD_CODE_RE = re.compile(r"Field Code:\s*([^;]+?)(?=\s*(?:;|Line Item|$))",
                           re.IGNORECASE)
LINE_EXT_RE = re.compile(r"Line Item External Id:\s*(\S+)", re.IGNORECASE)

IMPORTED_RE = re.compile(r"([\d,]+)\s+Purchase Orders? imported successfully",
                         re.IGNORECASE)
FAILED_RE = re.compile(r"([\d,]+)\s+Purchase Orders? failed to import",
                       re.IGNORECASE)


def flatten(message: str) -> str:
    """
    The message as one line of text.

    Concur writes its multi-part messages with literal `<br/>` and `</br>`
    tags in the cell, sometimes both, and escapes the odd entity. The markup
    is stripped for reading and for every regex below; `message` keeps what
    came, so the detail panel can still show it exactly.
    """
    text = MARKUP_RE.sub(" ", message or "")
    return SPACE_RE.sub(" ", html.unescape(text)).strip()


def parse_message(message: str) -> dict:
    """Pull the order number and the error parts out of one message."""
    text = flatten(message)
    padded = text + " "                       # so a trailing number still ends
    po = PO_RE.search(padded)
    code = CODE_RE.search(text)
    err = ERRTEXT_RE.search(text)
    lvl = FIELD_LEVEL_RE.search(text)
    fcode = FIELD_CODE_RE.search(text)
    ext = LINE_EXT_RE.search(text)
    return {
        "text": text,
        "stated_po": (po.group(1).strip() if po else ""),
        "error_code": (code.group(1).strip() if code else ""),
        "error_text": (err.group(1).strip(" .") if err else ""),
        "field_level": (lvl.group(1).strip(" ;") if lvl else ""),
        "field_code": (fcode.group(1).strip(" ;") if fcode else ""),
        "line_item_external_id": (ext.group(1).strip(" .;") if ext else ""),
    }


def finding_kind(row: dict) -> str:
    """
    A stable slug for one result row, so the Findings tab can group by it.

    The error code is used where Concur gave one; the two messages that carry
    no code are the halves of a rejected order and get a slug each, because
    "the sequence is invalid" and "so the order was not imported" are one
    fault reported twice and it helps to see that in the chip counts.
    """
    if row.get("error_code"):
        return "concur_" + str(row["error_code"])
    text = (row.get("text") or "").lower()
    if "sequence of the record types" in text:
        return "concur_sequence"
    if "was not imported" in text:
        return "concur_not_imported"
    return "concur_" + (row.get("level") or "error").lower()


def summary_numbers(rows: list[dict]) -> dict:
    """
    The counts Concur states in the Info row that closes the report.

    Worth keeping because they are the only independent check on this whole
    module: if the report says 31 orders failed and the rows resolve to 31
    orders, the line-number resolution is right.
    """
    out = {"imported": None, "failed": None}
    for r in rows:
        m = IMPORTED_RE.search(r["text"])
        if m:
            out["imported"] = int(m.group(1).replace(",", ""))
        m = FAILED_RE.search(r["text"])
        if m:
            out["failed"] = int(m.group(1).replace(",", ""))
    return out


# ------------------------------------------------------------------ loading


def is_results_file(path: Path) -> bool:
    """
    Whether a dropped file is a results report rather than an extract.

    A spreadsheet always is - nothing else this tool reads is one. A text file
    is only taken as a report if the headings are actually in it, so a CHAR-
    padded extract that happens to be saved as .csv is still an extract.
    """
    path = Path(path)
    if path.suffix.lower() in SPREADSHEET_SUFFIXES:
        return True
    if path.suffix.lower() not in TEXT_SUFFIXES:
        return False
    try:
        head = path.read_bytes()[:4096].decode("utf-8-sig", errors="replace").lower()
    except OSError:
        return False
    return COL_RECORD in head and COL_MESSAGE in head


def read_results(path: Path) -> list[dict]:
    """The report as rows, parsed but not yet resolved against any file."""
    grid = read_grid(Path(path))
    header_at, cols = find_header(grid)
    rows = []
    for i, raw in enumerate(grid[header_at + 1:], start=1):
        def cell(which: str) -> str:
            j = cols[which]
            if j < 0 or j >= len(raw):
                return ""
            v = raw[j]
            return "" if v is None else str(v).strip()
        message = cell("message")
        record_id = _as_int(raw[cols["record"]] if cols["record"] < len(raw) else None)
        level = cell("level") or "Error"
        if not message and record_id is None:
            continue                      # a blank row at the foot of the sheet
        rows.append({"row_no": i, "level": level.title(), "record_id": record_id,
                     "message": message, **parse_message(message)})
    if not rows:
        raise ResultsError("The report has headings but no rows under them.")
    return rows


def load_results(conn: sqlite3.Connection, path: Path,
                 commit: bool = True) -> dict:
    """
    Read one results report into the database, replacing whatever report was
    loaded before, and resolve it against the purchase order file that is
    loaded now.

    Replacing rather than accumulating matches how the three extract files are
    handled: one report is in play at a time, and it is the answer to one run.
    """
    path = Path(path)
    rows = read_results(path)
    numbers = summary_numbers(rows)
    levels: dict[str, int] = {}
    for r in rows:
        levels[r["level"]] = levels.get(r["level"], 0) + 1

    replace_file(conn, RESULTS_KIND)
    data = path.read_bytes()
    cur = conn.execute(
        "INSERT INTO ERP_Concur_Files (kind, file_name, file_path, sha256, "
        "byte_size, line_count, record_counts, newline, trailing_nl) "
        "VALUES (?,?,?,?,?,?,?,?,?)",
        (RESULTS_KIND, path.name, str(path), hashlib.sha256(data).hexdigest(),
         len(data), len(rows),
         json.dumps({**levels, "imported": numbers["imported"],
                     "failed": numbers["failed"]}), "", 0))
    file_id = cur.lastrowid

    conn.executemany(
        "INSERT INTO ERP_Concur_Results (file_id, row_no, level, record_id, "
        "message, text, error_code, error_text, field_level, field_code, "
        "line_item_external_id, stated_po) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
        [(file_id, r["row_no"], r["level"], r["record_id"], r["message"],
          r["text"], r["error_code"], r["error_text"], r["field_level"],
          r["field_code"], r["line_item_external_id"], r["stated_po"])
         for r in rows])

    resolution = resolve(conn, commit=False)
    if commit:
        conn.commit()
    result = {"kind": RESULTS_KIND, "label": RESULTS_LABEL, "file": path.name,
              "file_id": file_id, "rows": len(rows), "levels": levels,
              "run": run_label(path.name), **numbers, **resolution}
    log.info("%s: %s", path.name, result)
    return result


RUN_RE = re.compile(r"run[ _-]*(\d+)", re.IGNORECASE)


def run_label(file_name: str) -> str:
    """'Run-73' out of the file name, for the header pill and the manifest."""
    m = RUN_RE.search(file_name or "")
    return f"Run {m.group(1)}" if m else ""


# --------------------------------------------------------------- resolving


def line_index(conn: sqlite3.Connection) -> dict[int, tuple[str, str]]:
    """
    line number in the purchase order file -> (order number, what is there).

    Built from all three record tables at once, because the identifier can
    point at any of them: a 5001 error points at the header, and a sequence
    error points at the 210 that preceded it. Lines and addresses that never
    got a header (`po_key IS NULL`, a truncated file) are left out - there is
    no order to resolve them to, and that is worth reporting rather than
    guessing at.
    """
    index: dict[int, tuple[str, str]] = {}
    for sql, scope in (
        ("SELECT line_no, po_number FROM ERP_Concur_PoHeaders", "po"),
        ("SELECT l.line_no, h.po_number FROM ERP_Concur_PoLines l "
         "JOIN ERP_Concur_PoHeaders h ON h.po_key = l.po_key", "line"),
        ("SELECT a.line_no, h.po_number FROM ERP_Concur_PoAddresses a "
         "JOIN ERP_Concur_PoHeaders h ON h.po_key = a.po_key", "address"),
    ):
        for line_no, po_number in conn.execute(sql):
            index[int(line_no)] = (po_number, scope)
    return index


def resolve(conn: sqlite3.Connection, commit: bool = True) -> dict:
    """
    Point every row of the loaded report at a purchase order, and check it.

    Run on load and again after any file is loaded, because re-dropping the
    purchase order file changes what each line number means. Everything is
    recomputed from scratch, so it can never be half-updated.

    `aligned` is only set where both readings are available: the message named
    an order *and* the line number resolved to one. 0 means they disagree,
    which means the report and the purchase order file are from different
    runs - the numbers still add up and every one of them is wrong.
    """
    index = line_index(conn)
    rows = [dict(r) for r in conn.execute(
        "SELECT result_key, record_id, stated_po, level FROM ERP_Concur_Results")]
    updates = []
    ok = bad = 0
    for r in rows:
        po, scope, line_no, aligned = "", "", None, None
        hit = index.get(r["record_id"]) if r["record_id"] is not None else None
        if hit:
            po, scope = hit
            line_no = r["record_id"]
            if r["stated_po"]:
                aligned = 1 if po == r["stated_po"] else 0
                ok, bad = (ok + 1, bad) if aligned else (ok, bad + 1)
        updates.append((po, scope, line_no, aligned, r["result_key"]))
    conn.executemany(
        "UPDATE ERP_Concur_Results SET resolved_po = ?, resolved_scope = ?, "
        "resolved_line_no = ?, aligned = ? WHERE result_key = ?", updates)
    if commit:
        conn.commit()

    failed = failed_po_numbers(conn)
    stated_only = [r for r in rows
                   if r["level"] == "Error" and not index.get(r["record_id"])]
    return {"checked": ok + bad, "aligned": ok, "misaligned": bad,
            "failed_pos": len(failed),
            "unresolved_error_rows": len(
                [r for r in stated_only if not r["stated_po"]])}


def failed_po_numbers(conn: sqlite3.Connection) -> list[str]:
    """
    Every order the report says failed - the set the pick button acts on.

    The order the message names wins over the one the line number found; where
    a message named none, the line number is all there is. A row that gives
    neither is counted as unresolved and reported rather than dropped.
    """
    return [r[0] for r in conn.execute(
        "SELECT DISTINCT CASE WHEN stated_po <> '' THEN stated_po "
        "  ELSE resolved_po END AS po FROM ERP_Concur_Results "
        "WHERE level = 'Error' AND (stated_po <> '' OR resolved_po <> '') "
        "ORDER BY po")]


# ------------------------------------------------------------------ reading
# out again - what the page and the findings pass ask for.


EFFECTIVE_PO = ("CASE WHEN r.stated_po <> '' THEN r.stated_po "
                "ELSE r.resolved_po END")


def file_row(conn: sqlite3.Connection) -> dict | None:
    r = conn.execute(
        "SELECT * FROM ERP_Concur_Files WHERE kind = ? ORDER BY file_id DESC "
        "LIMIT 1", (RESULTS_KIND,)).fetchone()
    if r is None:
        return None
    row = dict(r)
    row["record_counts"] = json.loads(row.get("record_counts") or "{}")
    row["run"] = run_label(row["file_name"])
    return row


def summary(conn: sqlite3.Connection) -> dict:
    """Everything the Results tab needs above its table."""
    f = file_row(conn)
    if not f:
        return {"loaded": False}
    counts = conn.execute(
        "SELECT COUNT(*) n, "
        " SUM(level = 'Error') n_error, "
        " SUM(level = 'Error' AND resolved_po = '' AND stated_po = '') n_unresolved, "
        " SUM(aligned = 1) n_aligned, SUM(aligned = 0) n_misaligned "
        "FROM ERP_Concur_Results").fetchone()
    failed = failed_po_numbers(conn)
    in_file = {r[0] for r in conn.execute(
        "SELECT po_number FROM ERP_Concur_PoHeaders")}
    picked = {r[0] for r in conn.execute(
        "SELECT id FROM ERP_Concur_Selection WHERE kind = 'po'")}
    po_file = conn.execute(
        "SELECT file_name, line_count FROM ERP_Concur_Files WHERE kind = 'po'"
    ).fetchone()
    stated = f["record_counts"] or {}
    return {
        "loaded": True, "file": f["file_name"], "run": f["run"],
        "loaded_at": f["loaded_at"], "rows": counts["n"],
        "errors": counts["n_error"] or 0,
        "unresolved": counts["n_unresolved"] or 0,
        "aligned": counts["n_aligned"] or 0,
        "misaligned": counts["n_misaligned"] or 0,
        "stated_imported": stated.get("imported"),
        "stated_failed": stated.get("failed"),
        "failed_pos": len(failed),
        "failed_in_file": len([p for p in failed if p in in_file]),
        "failed_picked": len([p for p in failed if p in picked]),
        "po_file": (po_file["file_name"] if po_file else ""),
        "po_file_lines": (po_file["line_count"] if po_file else 0),
        "max_record_id": conn.execute(
            "SELECT MAX(record_id) FROM ERP_Concur_Results").fetchone()[0] or 0,
        "codes": [dict(r) for r in conn.execute(
            "SELECT CASE WHEN error_code <> '' THEN 'Error ' || error_code "
            "            WHEN text LIKE '%sequence of the record types%' "
            "              THEN 'Record sequence' "
            "            WHEN text LIKE '%was not imported%' "
            "              THEN 'Order not imported' "
            "            ELSE level END AS code, "
            "       level, COUNT(*) n, MIN(text) sample "
            "FROM ERP_Concur_Results GROUP BY code, level "
            "ORDER BY (level = 'Error') DESC, n DESC")],
    }


def failed_orders(conn: sqlite3.Connection) -> list[dict]:
    """One row per failed order, for the Orders view of the Results tab."""
    return [dict(r) for r in conn.execute(
        f"SELECT {EFFECTIVE_PO} AS po_number, COUNT(*) n_row, "
        "  GROUP_CONCAT(DISTINCT CASE WHEN r.error_code <> '' "
        "    THEN r.error_code END) codes, MIN(r.text) sample, "
        "  MIN(r.record_id) first_record, "
        "  (SELECT h.po_key FROM ERP_Concur_PoHeaders h "
        f"    WHERE h.po_number = {EFFECTIVE_PO} LIMIT 1) po_key, "
        "  (SELECT h.vendor_code FROM ERP_Concur_PoHeaders h "
        f"    WHERE h.po_number = {EFFECTIVE_PO} LIMIT 1) vendor_code, "
        f"  ({EFFECTIVE_PO} IN (SELECT id FROM ERP_Concur_Selection "
        "     WHERE kind = 'po')) picked "
        "FROM ERP_Concur_Results r "
        f"WHERE r.level = 'Error' AND {EFFECTIVE_PO} <> '' "
        "GROUP BY po_number ORDER BY po_number")]


# ------------------------------------------------------------------ findings


# One finding per error row is right for a run like the sample - 43 rows - and
# wrong for a run that rejected everything. Past this many, the rest are
# summarised in one finding instead of flooding the table.
FINDING_CAP = 800


def add_findings(f) -> None:
    """
    Add what Concur said to the findings pass, as `concur_*` kinds.

    Called from ERP_Concur_Findings.rebuild so that one table holds both
    halves: what this tool predicts the file will do, and what Concur replied
    when it did it. They are worth reading together - a 5001 predicted by
    `both_expense_and_account` and a 5001 that came back is the same fault
    confirmed, and a rejection with no finding in front of it is a check this
    tool is missing.
    """
    conn = f.conn
    if not conn.execute("SELECT 1 FROM ERP_Concur_Results LIMIT 1").fetchone():
        return
    report = file_row(conn) or {}
    name = report.get("file_name", "the report")
    run = report.get("run") or ""

    rows = [dict(r) for r in conn.execute(
        f"SELECT r.*, {EFFECTIVE_PO} AS po, "
        f" (SELECT h.po_key FROM ERP_Concur_PoHeaders h WHERE h.po_number = {EFFECTIVE_PO} "
        "   LIMIT 1) po_key, "
        f" (SELECT h.vendor_code FROM ERP_Concur_PoHeaders h WHERE h.po_number = {EFFECTIVE_PO} "
        "   LIMIT 1) vendor_code "
        "FROM ERP_Concur_Results r WHERE r.level <> 'Info' "
        "ORDER BY r.row_no LIMIT ?", (FINDING_CAP,))]
    for r in rows:
        severity = "error" if r["level"] == "Error" else "warning"
        where = (f"record {r['record_id']}" if r["record_id"] is not None
                 else f"row {r['row_no']}")
        f.add(severity, finding_kind(r), "po" if r["po_key"] else "file",
              r["po_key"], r["text"],
              label=(f"PO {r['po']}" if r["po"] else where),
              field=r["field_code"] or "", po_number=r["po"] or "",
              vendor_code=r["vendor_code"] or "")

    total_bad = conn.execute(
        "SELECT COUNT(*) FROM ERP_Concur_Results WHERE level <> 'Info'"
    ).fetchone()[0]
    if total_bad > FINDING_CAP:
        f.add("info", "concur_report_capped", "file", None,
              f"{name} holds {total_bad} rows that are not Info; the first "
              f"{FINDING_CAP} are listed here. The Results tab has them all.",
              label=run or "results report")

    s = summary(conn)
    if s.get("misaligned"):
        f.add("error", "concur_report_mismatch", "file", None,
              f"{s['misaligned']} of {s['aligned'] + s['misaligned']} messages "
              "that name a purchase order name a different one from the order "
              "at that line of the loaded purchase order file. The report and "
              f"the file are from different runs, so every order resolved from "
              f"a line number is wrong. Load the file {name} was the answer to.",
              label=run or "results report")
    if s.get("unresolved"):
        f.add("warning", "concur_unresolved_row", "file", None,
              f"{s['unresolved']} error row(s) name no purchase order and "
              "resolve to no line of the loaded purchase order file, so they "
              "cannot be picked. Load the purchase order file this report "
              "answers, or pick those orders by hand.",
              label=run or "results report")
    stated_failed = s.get("stated_failed")
    if stated_failed is not None and stated_failed != s.get("failed_pos"):
        f.add("warning", "concur_count_mismatch", "file", None,
              f"{name} says {stated_failed} purchase order(s) failed, and the "
              f"rows resolve to {s['failed_pos']}. One of the two readings of "
              "the report is incomplete; the summary line is the one to "
              "believe.", label=run or "results report")
