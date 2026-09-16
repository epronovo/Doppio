"""
ERP_Concur_Export - write the picked slice of the three files back out.

Every line that goes into an extract is the line that came in, re-emitted
verbatim from `raw`. Nothing is re-serialised from the parsed columns, nothing
is trimmed, no date is reformatted, no field is re-quoted, and the records come
out in the order they were in the source file - which for the purchase order
file means any 400 allocations immediately before the 300 line they belong to,
then 210, then 220, then the 200 header, header last, the way
ConcurExtracts/PurchaseOrder.dtsx writes it.

That is the whole contract of this module, and it is what makes the output
usable: the question these extracts answer is "what does Concur do with the
real file, cut down to two purchase orders", and an output that had quietly
fixed the padding or the dates would answer a different question. The line
endings of the source file are reproduced too, including whether it ended with
one.

The files are deliberately NOT named the way the dtsx names its own
(purchase_order_import_<entity>_.txt and friends). A partial file that looks
like the nightly one gets picked up as the nightly one - the same trap
ADP_Concur_Export sidesteps with its own naming - so these carry a label and
a timestamp (STEMS below is the base name for each), and a manifest is
written beside them, itself named `_Selection_`, saying what was in the cut.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import logging
import re
import sqlite3
import sys
from pathlib import Path

from ERP_Concur_Db import (
    DEFAULT_DB_PATH,
    connect,
    resolve_db_path,
    selection_add,
    selection_clear,
    selection_summary,
)

BASE_DIR = Path(__file__).parent.resolve()
DEFAULT_OUTPUT_DIR = BASE_DIR / "output"

log = logging.getLogger("ERP_Concur_Export")

NEWLINES = {"LF": "\n", "CRLF": "\r\n", "CR": "\r"}

# kind -> the stem of the file written for it.
STEMS = {"po": "POInvoice",
         "vendor": "vendor",
         "receipt": "purch_receipt"}


def _slug(label: str) -> str:
    s = re.sub(r"[^A-Za-z0-9]+", "_", (label or "").strip()).strip("_")
    return s[:40]


def _file_row(conn: sqlite3.Connection, kind: str) -> dict | None:
    r = conn.execute("SELECT * FROM ERP_Concur_Files WHERE kind = ?",
                     (kind,)).fetchone()
    return dict(r) if r else None


def _write(path: Path, lines: list[str], newline: str, trailing: bool) -> int:
    """Write the lines with the source file's endings, bytes unaltered."""
    text = newline.join(lines) + (newline if (lines and trailing) else "")
    path.parent.mkdir(parents=True, exist_ok=True)
    # newline="" so Python does not translate anything on the way out; the
    # endings are already in the text.
    with open(path, "w", encoding="utf-8", newline="") as fh:
        fh.write(text)
    return len(lines)


# ------------------------------------------------------------- the three cuts


def po_lines(conn: sqlite3.Connection) -> tuple[list[str], dict]:
    """
    Every record of every picked purchase order, in source order.

    One query over a union of the three record tables rather than a walk per
    header: ordering by the line number they came from reproduces the file's
    own grouping for free, and there is no way for it to drift from what the
    parser read.

    A synthesized allocation (see ERP_Concur_Parse.synthesize_allocations) has
    no line_no of its own in the file - it was never in the file - so it is
    stored under its parent line's line_no, tying with it. The CASE below
    breaks that tie the same way a real 400 sits in the file: immediately
    *before* the 300 it allocates, not after (confirmed against production
    data - see ERP_Concur_Parse._load_po_file). Every genuine parsed record
    has a line_no nothing else shares, so the tiebreaker changes nothing about
    a file that carries no synthesized rows.
    """
    rows = conn.execute(
        "SELECT line_no, raw, what FROM ("
        " SELECT line_no, raw, 'line' AS what FROM ERP_Concur_PoLines l "
        "  WHERE l.po_key IN (SELECT h.po_key FROM ERP_Concur_PoHeaders h "
        "    JOIN ERP_Concur_Selection s ON s.kind = 'po' AND s.id = h.po_number) "
        " UNION ALL "
        " SELECT line_no, raw, 'allocation' FROM ERP_Concur_PoLineAllocations a "
        "  WHERE a.po_key IN (SELECT h.po_key FROM ERP_Concur_PoHeaders h "
        "    JOIN ERP_Concur_Selection s ON s.kind = 'po' AND s.id = h.po_number) "
        " UNION ALL "
        " SELECT line_no, raw, 'address' FROM ERP_Concur_PoAddresses a "
        "  WHERE a.po_key IN (SELECT h.po_key FROM ERP_Concur_PoHeaders h "
        "    JOIN ERP_Concur_Selection s ON s.kind = 'po' AND s.id = h.po_number) "
        " UNION ALL "
        " SELECT line_no, raw, 'header' FROM ERP_Concur_PoHeaders h "
        "  JOIN ERP_Concur_Selection s ON s.kind = 'po' AND s.id = h.po_number"
        ") ORDER BY line_no, CASE what WHEN 'allocation' THEN 0 WHEN 'line' THEN 1 "
        "  WHEN 'address' THEN 2 ELSE 3 END").fetchall()
    counts = {"header": 0, "line": 0, "allocation": 0, "address": 0}
    for r in rows:
        counts[r["what"]] += 1
    return [r["raw"] for r in rows], counts


def vendor_lines(conn: sqlite3.Connection) -> tuple[list[str], dict]:
    """
    The 100 Import Settings record, then the picked vendors.

    The settings record is re-emitted from the source file when there was one.
    If the vendor file that was loaded had none, one is written from the values
    the dtsx hardcodes - record_type 100, error_threshold 0, and two empty
    fields - because Concur takes the settings record as the first line of the
    file and a vendor file without it is rejected outright.
    """
    settings = conn.execute(
        "SELECT raw FROM ERP_Concur_Settings ORDER BY settings_key LIMIT 1"
    ).fetchone()
    out = [settings["raw"] if settings else "100,0,,"]
    rows = conn.execute(
        "SELECT v.raw FROM ERP_Concur_Vendors v "
        " JOIN ERP_Concur_Selection s ON s.kind = 'vendor' AND s.id = v.vendor_code "
        "ORDER BY v.line_no").fetchall()
    out += [r["raw"] for r in rows]
    return out, {"settings": 1, "vendor": len(rows),
                 "settings_source": "file" if settings else "default"}


def receipt_lines(conn: sqlite3.Connection) -> tuple[list[str], dict]:
    """
    Receipts for the picked purchase orders, matched on Purchase Order Number.

    Matched on the order, not on the line: a receipt whose Line Item External
    ID resolves to nothing is exactly the error-1001 case worth reproducing, so
    it is kept in the cut and reported as a finding rather than filtered out.
    Receipts for an order that is not in this purchase order file at all are
    left out - there is nothing for them to attach to.
    """
    rows = conn.execute(
        "SELECT r.raw FROM ERP_Concur_Receipts r "
        " JOIN ERP_Concur_Selection s ON s.kind = 'po' AND s.id = r.po_number "
        "ORDER BY r.line_no").fetchall()
    return [r["raw"] for r in rows], {"receipt": len(rows)}


CUTS = {"po": po_lines, "vendor": vendor_lines, "receipt": receipt_lines}


# -------------------------------------------------------------------- writing


def write_extract(conn: sqlite3.Connection, out_dir: Path | None = None,
                  label: str = "", kinds: list[str] | None = None,
                  manifest: bool = True, commit: bool = True) -> dict:
    """
    Write one file per loaded kind for the current selection, and record it.

    A kind whose source file is not loaded is skipped rather than written
    empty - an empty vendor file is a file that deletes nothing and means
    nothing, and a zero-byte drop in a Concur pickup folder is a support call.
    """
    out_dir = Path(out_dir or DEFAULT_OUTPUT_DIR)
    kinds = kinds or list(STEMS)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    tag = f"_{_slug(label)}" if label else ""
    summary = selection_summary(conn)

    written, totals, skipped = [], {}, []
    for kind in kinds:
        src = _file_row(conn, kind)
        if not src:
            skipped.append({"kind": kind, "why": "no source file loaded"})
            continue
        lines, counts = CUTS[kind](conn)
        data_rows = sum(v for k, v in counts.items() if isinstance(v, int)
                        and k != "settings")
        if not data_rows:
            skipped.append({"kind": kind, "why": "nothing picked resolves to a "
                                                 "record in this file"})
            continue
        name = f"{STEMS[kind]}{tag}_{stamp}.txt"
        path = out_dir / name
        n = _write(path, lines, NEWLINES.get(src["newline"] or "LF", "\n"),
                   bool(src["trailing_nl"]))
        totals.update({f"{kind}_{k}": v for k, v in counts.items()})
        written.append({"kind": kind, "name": name, "path": str(path),
                        "rows": n, "counts": counts,
                        "source": src["file_name"], "newline": src["newline"]})
        log.info("%s: %d lines", path, n)

    scope = (f"{summary['vendors']} vendor(s), {summary['pos']} purchase order(s)"
             if written else "nothing written")
    cur = conn.execute(
        "INSERT INTO ERP_Concur_Extracts (stamp, label, scope, n_vendor, n_po, "
        "n_line, n_allocation, n_address, n_receipt, files) "
        "VALUES (?,?,?,?,?,?,?,?,?,?)",
        (stamp, label, scope, totals.get("vendor_vendor", 0),
         totals.get("po_header", 0), totals.get("po_line", 0),
         totals.get("po_allocation", 0), totals.get("po_address", 0),
         totals.get("receipt_receipt", 0), json.dumps(written)))
    extract_key = cur.lastrowid

    manifest_file = None
    if manifest and written:
        manifest_file = write_manifest(conn, out_dir, stamp, tag, extract_key,
                                       written, skipped, summary)
    if commit:
        conn.commit()
    return {"extract_key": extract_key, "stamp": stamp, "label": label,
            "out_dir": str(out_dir), "files": written, "skipped": skipped,
            "manifest": manifest_file, "selection": summary}


def write_manifest(conn, out_dir: Path, stamp: str, tag: str, extract_key: int,
                   written: list[dict], skipped: list[dict],
                   summary: dict) -> str:
    """
    A plain-text note beside the files saying what is in them.

    Worth the few lines: these extracts get mailed to somebody who loads them
    a week later, and "which POs were in the one that failed" is otherwise a
    question only this database can answer.
    """
    vendors = [dict(r) for r in conn.execute(
        "SELECT s.id, s.direct, s.reason, "
        " (SELECT vendor_name FROM ERP_Concur_Vendors v WHERE v.vendor_code = s.id "
        "  LIMIT 1) AS name "
        "FROM ERP_Concur_Selection s WHERE s.kind = 'vendor' ORDER BY s.id")]
    pos = [dict(r) for r in conn.execute(
        "SELECT s.id, s.direct, s.reason, h.vendor_code, h.order_date, "
        " (SELECT COUNT(*) FROM ERP_Concur_PoLines l WHERE l.po_key = h.po_key) n_line, "
        " (SELECT COUNT(*) FROM ERP_Concur_Receipts r WHERE r.po_number = s.id) n_rec "
        "FROM ERP_Concur_Selection s "
        "LEFT JOIN ERP_Concur_PoHeaders h ON h.po_number = s.id "
        "WHERE s.kind = 'po' ORDER BY s.id")]
    # Only the findings that describe something in this cut: a padded postal
    # code on a vendor that was not picked is not a finding about this file.
    # Findings carrying neither key (a missing file, a truncated tail) are
    # about the extract as a whole and always count.
    findings = [dict(r) for r in conn.execute(
        "SELECT severity, kind, COUNT(*) n FROM ERP_Concur_Findings f "
        "WHERE (f.po_number = '' AND f.vendor_code = '') "
        "   OR f.po_number IN (SELECT id FROM ERP_Concur_Selection WHERE kind = 'po') "
        "   OR f.vendor_code IN (SELECT id FROM ERP_Concur_Selection WHERE kind = 'vendor') "
        "GROUP BY severity, kind ORDER BY severity, kind")]

    synthetic_allocs = conn.execute(
        "SELECT COUNT(*) FROM ERP_Concur_PoLineAllocations a "
        "WHERE a.synthetic = 1 AND a.po_key IN "
        "(SELECT h.po_key FROM ERP_Concur_PoHeaders h "
        " JOIN ERP_Concur_Selection s ON s.kind = 'po' AND s.id = h.po_number)"
    ).fetchone()[0]

    L = [f"ERP -> Concur selection extract {extract_key}",
         f"Written {dt.datetime.now():%Y-%m-%d %H:%M:%S} (stamp {stamp})",
         "",
         "These are a SUBSET of the files ConcurExtracts/PurchaseOrder.dtsx",
         "produced, cut to the vendors and purchase orders listed below. Every",
         "line is byte-for-byte the line from the source file, in source order -",
         "with ONE exception: a 400 line allocation added by 'Fill missing",
         "allocations' is not from the source file at all; see below.",
         "Each carries a label and a timestamp, unlike the nightly",
         "purchase_order_import_* drop, so a partial file cannot be mistaken",
         "for it - see this manifest, named _Selection_ itself, for what was",
         "actually in the cut.",
         "",
         "FILES"]
    for w in written:
        L.append(f"  {w['name']}")
        L.append(f"      {w['rows']} line(s) from {w['source']} "
                 f"({w['newline']} line endings)  "
                 + ", ".join(f"{k}={v}" for k, v in w["counts"].items()))
    for s in skipped:
        L.append(f"  (no {s['kind']} file written - {s['why']})")
    if synthetic_allocs:
        L.append(f"      ⚠ {synthetic_allocs} of the allocation record(s) above "
                 "are SYNTHESIZED - a trivial Quantity x Unit Price allocation "
                 "this tool added because the line had none, not a real "
                 "account split read from any ERP. See ERP_Concur_Parse."
                 "synthesize_allocations for why one was added and what it "
                 "does and does not claim.")

    L += ["", f"VENDORS ({len(vendors)})"]
    for v in vendors:
        how = "picked" if v["direct"] else f"via {v['reason']}"
        L.append(f"  {v['id']:<14} {(v['name'] or '(not in the vendor file)'):<34} {how}")
    L += ["", f"PURCHASE ORDERS ({len(pos)})"]
    for p in pos:
        how = "picked" if p["direct"] else f"via {p['reason']}"
        if p["vendor_code"] is None:
            L.append(f"  {p['id']:<14} NOT IN THE PURCHASE ORDER FILE - "
                     f"nothing written for it ({how})")
        else:
            L.append(f"  {p['id']:<14} vendor {p['vendor_code']:<12} "
                     f"{p['order_date'] or '':<12} {p['n_line']} line(s), "
                     f"{p['n_rec']} receipt(s)  {how}")
    if summary["unresolved_vendors"] or summary["unresolved_pos"]:
        L += ["", "PICKED BUT NOT IN THE FILES THAT ARE LOADED"]
        for i in summary["unresolved_vendors"]:
            L.append(f"  vendor {i}")
        for i in summary["unresolved_pos"]:
            L.append(f"  PO {i}")
    if findings:
        L += ["", "FINDINGS COVERING THIS CUT"]
        for f in findings:
            L.append(f"  {f['severity']:<8} {f['kind']:<28} {f['n']}")
        L.append("  (nothing was withheld because of these - see the Findings "
                 "tab for the detail)")
    name = f"FMG_Concur_Selection{tag}_{stamp}_manifest.txt"
    (out_dir / name).write_text("\n".join(L) + "\n", encoding="utf-8")
    return name


def verify(conn: sqlite3.Connection, files: list[dict]) -> dict:
    """
    Prove the claim this module makes: every line written is a line that was
    read, unchanged - with one declared exception.

    Each emitted line is looked up in the raw text the parser stored. Anything
    the database does not recognise is reported. Run by the test harness and by
    the Extract tab's Verify button - it is cheap, and "the subset is the
    source" is the one property worth checking every time rather than trusting.

    A 400 line added by synthesize_allocations is, on purpose, not a line that
    was read - it is counted separately here (`synthetic`) rather than folded
    into `unknown`, so a synthesized allocation never looks like the parser
    failure `unknown` exists to catch, but also never quietly passes as "came
    from the file" the way everything else in `checked` does.
    """
    known = {r[0] for r in conn.execute(
        "SELECT raw FROM ERP_Concur_PoLines UNION "
        "SELECT raw FROM ERP_Concur_PoLineAllocations WHERE synthetic = 0 UNION "
        "SELECT raw FROM ERP_Concur_PoAddresses UNION "
        "SELECT raw FROM ERP_Concur_PoHeaders UNION "
        "SELECT raw FROM ERP_Concur_Vendors UNION "
        "SELECT raw FROM ERP_Concur_Settings UNION "
        "SELECT raw FROM ERP_Concur_Receipts")}
    known.add("100,0,,")                      # the default settings record
    synthetic_raws = {r[0] for r in conn.execute(
        "SELECT raw FROM ERP_Concur_PoLineAllocations WHERE synthetic = 1")}
    checked = unknown = synthetic = 0
    strangers = []
    for w in files:
        for line in Path(w["path"]).read_text(encoding="utf-8").splitlines():
            checked += 1
            if line in synthetic_raws:
                synthetic += 1
            elif line not in known:
                unknown += 1
                if len(strangers) < 5:
                    strangers.append({"file": w["name"], "line": line[:120]})
    return {"checked": checked, "unknown": unknown, "synthetic": synthetic,
            "strangers": strangers, "ok": unknown == 0}


# ------------------------------------------------------------------------ CLI


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Write a Concur extract for a set of vendors or purchase "
                    "orders. With no --vendor/--po the current selection (the "
                    "one the app shows) is used as it stands.")
    ap.add_argument("--vendor", action="append", default=[], metavar="CODE",
                    help="vendor code to pick; repeatable. Pulls in that "
                         "vendor's orders and their receipts.")
    ap.add_argument("--po", action="append", default=[], metavar="NUMBER",
                    help="purchase order number to pick; repeatable. Pulls in "
                         "the vendor behind it.")
    ap.add_argument("--replace", action="store_true",
                    help="clear the app's selection first, so only what is "
                         "named here is written")
    ap.add_argument("--label", default="", help="goes in the file names")
    ap.add_argument("--out", default=str(DEFAULT_OUTPUT_DIR))
    ap.add_argument("--kind", action="append", default=[],
                    choices=list(STEMS), help="write only these files")
    ap.add_argument("--db", default=None,
                    help=f"SQLite path (default {DEFAULT_DB_PATH})")
    ap.add_argument("--verify", action="store_true",
                    help="check every written line against the source")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(message)s",
                        handlers=[logging.StreamHandler(sys.stdout)])
    conn = connect(args.db)
    print(f"Database : {resolve_db_path(args.db)}")
    if args.replace:
        selection_clear(conn)
    if args.vendor or args.po:
        selection_add(conn, vendors=args.vendor, pos=args.po, reason="CLI")
    result = write_extract(conn, out_dir=Path(args.out), label=args.label,
                           kinds=args.kind or None)
    if not result["files"]:
        print("Nothing written. " + "; ".join(
            f"{s['kind']}: {s['why']}" for s in result["skipped"]))
        return 1
    for w in result["files"]:
        print(f"  {w['path']}  {w['rows']} line(s)")
    if result["manifest"]:
        print(f"  {Path(args.out) / result['manifest']}")
    if args.verify:
        v = verify(conn, result["files"])
        print(f"Verify   : {v['checked']} line(s) checked, "
              f"{v['unknown']} not found in the source"
              + (f" ({v['synthetic']} synthesized allocation(s), by design)"
                 if v["synthetic"] else "")
              + f" -> {'OK' if v['ok'] else 'FAILED'}")
        return 0 if v["ok"] else 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
