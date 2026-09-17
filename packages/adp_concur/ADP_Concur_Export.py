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

from ADP_Concur_Db import (DEFAULT_DB_PATH, RECORD_TYPES, SOURCES,
                           connect, resolve_db_path)
from ADP_Concur_Map import (
    build_import_settings,
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

    There is no per-roster suffix any more. The US / non-US split is gone and
    both systems of record go into one file, so there is only ever one.
    """
    when = when or datetime.now()
    extract = cfg.get("extract") or {}
    pattern = (extract.get("selection_file_name")
               or "FMG_Concur_Employee_Selection_{stamp}.txt") if selection else (
        extract.get("file_name") or "FMG_Concur_Employee_{stamp}.txt")
    return pattern.format(stamp=when.strftime("%Y%m%d_%H%M%S"),
                          date=when.strftime("%Y%m%d"),
                          time=when.strftime("%H%M%S"))


def record_types_for(cfg: dict) -> list[str]:
    """
    Every record type this file can carry, in writing order.

    The union across the sources, because one file carries them all now - which
    source writes which is decided per person, in selected_employees(). In
    config because it is a real question rather than a constant: the July file
    that loaded carried 305 and 360 only, and every 350 Concur has seen since
    has been rejected on its Travel Class Name.
    """
    by_source = cfg.get("records_by_source") or {}
    wanted = {t for types in by_source.values() for t in types} or set(RECORD_TYPES)
    return [t for t in RECORD_TYPES if t in wanted]


def sort_by_hierarchy(people: list[dict]) -> list[dict]:
    """
    Active first, then down the supervisor tree - approvers before reports.

    Concur resolves an approver against what it already holds, so a record
    naming somebody who has not been created yet loads without them. Writing
    the tree top-down removes the problem for everybody inside the file at
    once, which is what let the two roster files become one: an approver in
    UKG is simply written above the ADP person pointing at them.

    Active before inactive is the outer key, as asked. It does mean an active
    person can be written above their own inactive manager - but that pairing
    is already broken (Concur will not resolve an approver who is inactive) and
    the sort should not hide it by burying the leaver in the middle of the
    tree.

    Everyone whose supervisor is not in this list is a root: the top of the
    chain, somebody reporting outside the load, and anybody left over from a
    cycle. Cycles cannot be ordered - that is what a cycle means - so whatever
    the walk has not reached is appended in name order rather than dropped.
    Losing a record to keep a sort tidy would be the worse bug by far.
    """
    by_fn: dict[str, dict] = {}
    for e in people:
        fn = str(e.get("file_number") or "")
        if fn:
            by_fn.setdefault(fn, e)

    def name_key(e: dict) -> tuple:
        return (str(e.get("legal_last_name") or ""),
                str(e.get("legal_first_name") or ""),
                str(e.get("file_number") or ""))

    reports: dict[str, list[dict]] = {}
    roots: list[dict] = []
    for e in people:
        sup = str(e.get("supervisor_id") or "")
        if sup and sup in by_fn and sup != str(e.get("file_number") or ""):
            reports.setdefault(sup, []).append(e)
        else:
            roots.append(e)

    out: list[dict] = []
    seen: set[int] = set()

    def walk(group: list[dict]) -> None:
        # Iterative rather than recursive: a deep chain in a company this size
        # is fine either way, but a cycle that slipped through would be a stack
        # overflow instead of a handled case.
        stack = sorted(group, key=name_key, reverse=True)
        while stack:
            e = stack.pop()
            if id(e) in seen:
                continue
            seen.add(id(e))
            out.append(e)
            kids = reports.get(str(e.get("file_number") or ""), [])
            stack.extend(sorted((k for k in kids if id(k) not in seen),
                                key=name_key, reverse=True))

    walk(roots)

    # Whatever the walk never reached - people inside a supervisor cycle.
    left = [e for e in people if id(e) not in seen]
    out.extend(sorted(left, key=name_key))

    # Status is the outer key, so the whole tree is walked first and then split
    # in place. The partition is stable, which is what keeps the hierarchy
    # inside each half: an active approver still precedes their active reports,
    # and an inactive one still precedes their inactive reports.
    #
    # The one pairing it cannot honour is an active person reporting to an
    # inactive one - the approver lands in the second half, below them. That is
    # the instruction doing what it says, and it is not hiding anything: an
    # approver who has left is a problem the Exceptions tab already raises, and
    # burying the leaver mid-tree to make the file look ordered would only make
    # it harder to see.
    active = [e for e in out if str(e.get("concur_status") or "") == "Y"]
    inactive = [e for e in out if str(e.get("concur_status") or "") != "Y"]
    return active + inactive


def collect(conn: sqlite3.Connection, cfg: dict,
            keys: list[int] | None = None) -> tuple[list[list[str]], dict]:
    """
    Every record the extract will carry, in the configured order.

    Returns the lines and a per-type count. Built here rather than in
    ADP_Concur_Map.build_records() because the ordering is an export concern.

    The file opens with the 100 Import Settings record, which SAP requires once
    per file and which is not per-employee - so it is prepended here rather
    than living in a field map with the rest.

    One file carries both sources. Which record types a person produces is
    decided by the system they came from - UKG has no 350 tab - and that is
    applied inside selected_employees() rather than here.
    """
    cfg_extract = cfg.get("extract") or {}
    types = record_types_for(cfg)
    widths = {rt: layout_width(conn, rt) for rt in types}
    people = {rt: selected_employees(conn, rt, cfg, keys) for rt in types}

    # The 305s carry the hierarchy, so they are the ones that are sorted. The
    # other three follow the 305 order for the people they cover, so the whole
    # file reads the same way down.
    ordered_305 = sort_by_hierarchy(people.get("305", []))
    rank = {str(e.get("file_number")): i for i, e in enumerate(ordered_305)}
    people["305"] = ordered_305
    for rt in types:
        if rt == "305":
            continue
        people[rt] = sorted(
            people[rt],
            key=lambda e: (rank.get(str(e.get("file_number")), len(rank)),
                           str(e.get("file_number") or "")))

    counts = {rt: len(people.get(rt, ())) for rt in RECORD_TYPES}

    lines: list[list[str]] = [build_import_settings(cfg)]
    if cfg_extract.get("order") == "by_employee":
        # Keyed on file number so a person's records stay together; the 305
        # order is the one that drives the file.
        wanted = {rt: {e["file_number"]: e for e in people[rt]} for rt in people}
        seen = []
        for rt in types:
            for e in people[rt]:
                if e["file_number"] not in seen:
                    seen.append(e["file_number"])
        for fn in seen:
            for rt in types:
                emp = wanted[rt].get(fn)
                if emp is not None:
                    lines.append(build_record(emp, rt, widths[rt], cfg))
    else:
        for rt in types:
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
              "n_360": counts["360"], "n_700": counts["700"],
              "bytes": len(text.encode(
                  (cfg.get("extract") or {}).get("encoding", "utf-8"))),
              "dry_run": dry_run, "scope": scope,
              "record_types": record_types_for(cfg),
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


def ADP_Concur_export_all(conn: sqlite3.Connection, cfg: dict | None = None,
                          out_dir: str | Path | None = None,
                          dry_run: bool = False) -> list[dict]:
    """
    Write the extract. One file, which is the whole point of this version.

    Kept as a list of one so every caller that walked the old two-file result
    still works, and so the front end's "write both files" button needed no
    special case on the day the split went away.
    """
    return [ADP_Concur_export(conn, cfg, out_dir=out_dir, dry_run=dry_run)]


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
    ap.add_argument("--all-rosters", "--all", action="store_true",
                    dest="all_rosters",
                    help="Kept for the old command line; there is one file now")
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
    print(f"{verb} {res['records']} record(s): 1 x 100, {res['n_305']} x 305, "
          f"{res['n_350']} x 350, {res['n_360']} x 360, {res['n_700']} x 700")
    print(f"  {res['path']}  ({res['bytes']:,} bytes)")
    conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
