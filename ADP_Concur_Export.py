"""
ADP_Concur_Export - write the flat file SAP Concur picks up.

One file carries all three record types. Each line is one record: the type in
the first field, then the template's columns in order, every position present
whether or not it has a value - which is what "All fields must be represented"
on the 350 and 360 tabs means.

Two orderings are offered because sites differ on this. 'by_type' writes every
305, then every 350, then every 360, which is the safer default: Concur creates
the profile from the 305, and the 350 and 360 records attach to a profile that
has to exist already. 'by_employee' writes a person's three records together.

Nothing here decides who is in the file - ADP_Concur_Map.selected_employees()
does, from the include flags, the configured scope and the exception list.
"""
from __future__ import annotations

import argparse
import csv
import io
import sqlite3
from datetime import datetime
from pathlib import Path

from ADP_Concur_Db import DEFAULT_DB_PATH, connect, resolve_db_path
from ADP_Concur_Map import (
    build_record,
    layout_width,
    load_config,
    selected_employees,
)

BASE_DIR = Path(__file__).parent.resolve()
DEFAULT_OUTPUT_DIR = BASE_DIR / "output" / "adp_concur"

QUOTING = {"minimal": csv.QUOTE_MINIMAL, "all": csv.QUOTE_ALL, "none": csv.QUOTE_NONE}


def outbound_dir(cfg: dict) -> Path:
    """Where the file lands - the configured folder, or output/adp_concur/."""
    configured = (cfg.get("extract") or {}).get("outbound_dir") or ""
    return Path(configured).expanduser() if configured else DEFAULT_OUTPUT_DIR


def extract_file_name(cfg: dict, when: datetime | None = None,
                      selection: bool = False) -> str:
    """
    {stamp} in the configured name becomes yyyymmdd_HHMMSS.

    A selection gets its own pattern, because the outbound folder is a Concur
    pickup and a partial file that looks exactly like a full one is the kind of
    thing that gets loaded by accident at four in the afternoon.
    """
    when = when or datetime.now()
    extract = cfg.get("extract") or {}
    pattern = (extract.get("selection_file_name")
               or "FMG_Concur_Employee_Selection_{stamp}.txt") if selection else (
        extract.get("file_name") or "FMG_Concur_Employee_{stamp}.txt")
    return pattern.format(stamp=when.strftime("%Y%m%d_%H%M%S"),
                          date=when.strftime("%Y%m%d"),
                          time=when.strftime("%H%M%S"))


def collect(conn: sqlite3.Connection, cfg: dict,
            keys: list[int] | None = None) -> tuple[list[list[str]], dict]:
    """
    Every record the extract will carry, in the configured order.

    Returns the lines and a per-type count. Built here rather than in
    ADP_Concur_Map.build_records() because the ordering is an export concern.
    """
    cfg_extract = cfg.get("extract") or {}
    widths = {rt: layout_width(conn, rt) for rt in ("305", "350", "360")}
    people = {rt: selected_employees(conn, rt, cfg, keys)
              for rt in ("305", "350", "360")}
    counts = {rt: len(people[rt]) for rt in people}

    lines: list[list[str]] = []
    if cfg_extract.get("order") == "by_employee":
        # Keyed on file number so the three records for one person stay
        # together; the 305 order is the one that drives the file.
        wanted = {rt: {e["file_number"]: e for e in people[rt]} for rt in people}
        seen = []
        for rt in ("305", "350", "360"):
            for e in people[rt]:
                if e["file_number"] not in seen:
                    seen.append(e["file_number"])
        for fn in seen:
            for rt in ("305", "350", "360"):
                emp = wanted[rt].get(fn)
                if emp is not None:
                    lines.append(build_record(emp, rt, widths[rt], cfg))
    else:
        for rt in ("305", "350", "360"):
            for emp in people[rt]:
                lines.append(build_record(emp, rt, widths[rt], cfg))

    return lines, counts


def render(lines: list[list[str]], cfg: dict) -> str:
    """The file as text, so it can be previewed without being written."""
    extract = cfg.get("extract") or {}
    buf = io.StringIO()
    writer = csv.writer(
        buf,
        delimiter=extract.get("delimiter", ","),
        quoting=QUOTING.get(extract.get("quote", "minimal"), csv.QUOTE_MINIMAL),
        quotechar='"',
        escapechar="\\" if extract.get("quote") == "none" else None,
        lineterminator=extract.get("line_ending", "\r\n"),
    )
    writer.writerows(lines)
    return buf.getvalue()


def ADP_Concur_export(conn: sqlite3.Connection, cfg: dict | None = None,
                      out_dir: str | Path | None = None,
                      file_name: str | None = None,
                      keys: list[int] | None = None,
                      selection_label: str = "",
                      dry_run: bool = False) -> dict:
    """
    Write the extract.

    `keys` narrows it to a chosen set of people - one manager's organisation,
    a filtered list, whatever was ticked. The file is named and recorded as a
    selection so a pilot load is never mistaken for the full company, and
    `selection_label` is stored alongside it saying what was chosen.

    dry_run builds everything and reports the counts without touching the
    disk, which is what the front end previews with.
    """
    cfg = cfg or load_config()
    lines, counts = collect(conn, cfg, keys)
    text = render(lines, cfg)

    directory = Path(out_dir).expanduser() if out_dir else outbound_dir(cfg)
    name = file_name or extract_file_name(cfg, selection=keys is not None)
    target = directory / name
    scope = ",".join(f"{k}={v}" for k, v in (cfg.get("scope") or {}).items())
    if keys is not None:
        scope = (f"selection of {len(keys)}"
                 + (f" ({selection_label})" if selection_label else "")
                 + "; " + scope)

    result = {"file_name": name, "path": str(target), "records": len(lines),
              "n_305": counts["305"], "n_350": counts["350"],
              "n_360": counts["360"], "bytes": len(text.encode(
                  (cfg.get("extract") or {}).get("encoding", "utf-8"))),
              "dry_run": dry_run, "scope": scope,
              "selection": keys is not None,
              "selected": len(keys) if keys is not None else None,
              "selection_label": selection_label,
              "preview": text.splitlines()[:5]}
    if dry_run:
        return result

    directory.mkdir(parents=True, exist_ok=True)
    encoding = (cfg.get("extract") or {}).get("encoding", "utf-8")
    with open(target, "w", encoding=encoding, newline="") as fh:
        fh.write(text)

    conn.execute(
        "INSERT INTO ADP_Concur_Extracts "
        "(file_name, file_path, n_305, n_350, n_360, scope, delimiter) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)",
        (name, str(target), counts["305"], counts["350"], counts["360"], scope,
         (cfg.get("extract") or {}).get("delimiter", ",")))
    conn.commit()
    return result


def held_back(conn: sqlite3.Connection,
              keys: list[int] | None = None) -> list[dict]:
    """
    Who the extract is leaving out, and why.

    Anyone with a blocking error, plus anyone excluded by an include flag.
    Shown beside the export button so the file is never quietly short. With a
    selection, only people inside that selection are reported - the other
    hundred are not being "held back", they were simply not chosen.
    """
    scope, args = "", []
    if keys is not None:
        if not keys:
            return []
        scope = f" AND e.employee_key IN ({','.join('?' * len(keys))})"
        args = list(keys)
    return [dict(r) for r in conn.execute(
        f"""
        SELECT e.file_number,
               e.legal_last_name || ', ' || e.legal_first_name AS name,
               e.position_status, e.concur_status,
               e.include_305, e.include_350, e.include_360,
               (SELECT GROUP_CONCAT(x.message, ' | ') FROM ADP_Concur_Exceptions x
                 WHERE x.employee_key = e.employee_key AND x.severity = 'error')
                   AS reasons
          FROM ADP_Concur_Employees e
         WHERE e.row_state <> 'deleted'
           AND (EXISTS (SELECT 1 FROM ADP_Concur_Exceptions x
                         WHERE x.employee_key = e.employee_key AND x.severity = 'error')
                OR e.include_305 = 0 OR e.include_350 = 0 OR e.include_360 = 0)
               {scope}
         ORDER BY e.legal_last_name, e.legal_first_name
        """, args)]


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Write the Concur employee extract.")
    ap.add_argument("--db", default=None, help=f"SQLite path (default {DEFAULT_DB_PATH})")
    ap.add_argument("--out-dir", default=None,
                    help=f"Where to write (default {DEFAULT_OUTPUT_DIR})")
    ap.add_argument("--name", default=None, help="Override the file name")
    ap.add_argument("--dry-run", action="store_true",
                    help="Report what would be written without writing it")
    ap.add_argument("--held-back", action="store_true",
                    help="List the people the extract leaves out and why")
    ap.add_argument("--under", metavar="FILE_NUMBER", action="append", default=[],
                    help="Scope the file to this person and everyone under "
                         "them. Repeatable - the perfect pilot load")
    ap.add_argument("--only", metavar="FILE_NUMBER", action="append", default=[],
                    help="Scope the file to these people exactly. Repeatable")
    args = ap.parse_args(argv)

    conn = connect(args.db)
    print(f"Database: {resolve_db_path(args.db)}")

    keys = None
    label = ""
    if args.under or args.only:
        from ADP_Concur_Hierarchy import ADP_Concur_subtree_keys
        keys = ADP_Concur_subtree_keys(conn, args.under) if args.under else []
        if args.only:
            marks = ",".join("?" * len(args.only))
            keys += [r[0] for r in conn.execute(
                f"SELECT employee_key FROM ADP_Concur_Employees "
                f"WHERE file_number IN ({marks}) AND row_state <> 'deleted'",
                args.only) if r[0] not in keys]
        label = "; ".join(filter(None, [
            "under " + ", ".join(args.under) if args.under else "",
            "only " + ", ".join(args.only) if args.only else ""]))
        print(f"Selection: {label} — {len(keys)} employee(s)")

    if args.held_back:
        rows = held_back(conn, keys)
        print(f"{len(rows)} employee(s) held back")
        for r in rows:
            flags = "".join(t for t, k in (("305", "include_305"), ("350", "include_350"),
                                           ("360", "include_360")) if not r[k])
            print(f"  {r['file_number']:<8} {(r['name'] or ''):<30} "
                  f"{r['reasons'] or ('excluded from ' + flags)}")
        conn.close()
        return 0

    res = ADP_Concur_export(conn, out_dir=args.out_dir, file_name=args.name,
                            keys=keys, selection_label=label,
                            dry_run=args.dry_run)
    verb = "Would write" if args.dry_run else "Wrote"
    print(f"{verb} {res['records']} record(s): {res['n_305']} x 305, "
          f"{res['n_350']} x 350, {res['n_360']} x 360")
    print(f"  {res['path']}  ({res['bytes']:,} bytes)")
    conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
