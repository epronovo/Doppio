"""
ADP_Concur_Hierarchy - the supervisor chain, read as a tree.

Every employee carries one derived `supervisor_id`, which is another
employee's `file_number`. That single link is the whole hierarchy, and it is
what Concur builds its approval routing from: the 305's Expense Report
Approver and both of the 360's approver fields are that number. So a link that
points at nobody is not a cosmetic problem - it is an approval chain that ends
in mid-air, and Concur rejects the record.

Nothing here is stored. The tree is derived from `supervisor_id` on demand,
because the maps can change under it at any moment and a cached tree would be
wrong the instant someone edits the Supervisor Map.

The recursive walks are written with a `path` column and a depth ceiling, so a
loop in the data comes back as a loop rather than hanging the query.
"""
from __future__ import annotations

import argparse
import re
import sqlite3

from ADP_Concur_Db import DEFAULT_DB_PATH, connect, resolve_db_path

# How far a walk will follow the chain before deciding it is not going to end.
# Real hierarchies are a handful of levels; anything past this is a loop or a
# data accident, and either way it is the checker's problem, not the walker's.
MAX_DEPTH = 40

# Enough of an employee to render a node without a second query.
NODE_COLUMNS = """
    e.employee_key, e.file_number, e.legal_first_name, e.legal_last_name,
    e.job_title, e.business_unit_desc, e.org_unit_1, e.org_unit_2,
    e.supervisor_id, e.supervisor_id_raw, e.reports_to_legal_name,
    e.position_status, e.concur_status, e.source, e.row_state,
    e.include_305, e.include_350, e.include_360
"""

LIVE = "e.row_state <> 'deleted'"


def _name(row: dict) -> str:
    return f"{row.get('legal_last_name') or ''}, {row.get('legal_first_name') or ''}".strip(", ")


def _decorate(rows: list[dict]) -> list[dict]:
    for r in rows:
        r["name"] = _name(r)
    return rows


# ------------------------------------------------------------------ upward


def ADP_Concur_chain_up(conn: sqlite3.Connection, file_number: str) -> dict:
    """
    The complete chain from one employee up to the top.

    Returns the chain ordered from the employee outwards - index 0 is the
    person, index 1 their supervisor, and so on - together with how the chain
    ended. `ended` is one of:

      top      the last person has no supervisor at all, which is the answer
               the Supervisor Map's 'Top of food chain' note describes
      broken   the last person names a supervisor no employee in this load has
      cycle    the chain came back to somebody it had already visited
      deep     it ran past MAX_DEPTH without resolving

    'broken' and 'cycle' are the two that matter: both mean Concur is being
    handed an approver it cannot resolve.
    """
    rows = [dict(r) for r in conn.execute(
        f"""
        WITH RECURSIVE up(employee_key, depth, path) AS (
            SELECT e.employee_key, 0, '/' || e.file_number || '/'
              FROM ADP_Concur_Employees e
             WHERE e.file_number = ? AND {LIVE}
            UNION ALL
            SELECT s.employee_key, up.depth + 1, up.path || s.file_number || '/'
              FROM up
              JOIN ADP_Concur_Employees e ON e.employee_key = up.employee_key
              JOIN ADP_Concur_Employees s ON s.file_number = e.supervisor_id
             WHERE e.supervisor_id <> ''
               AND s.row_state <> 'deleted'
               AND up.path NOT LIKE '%/' || s.file_number || '/%'
               AND up.depth < ?
        )
        SELECT {NODE_COLUMNS}, up.depth
          FROM up JOIN ADP_Concur_Employees e ON e.employee_key = up.employee_key
         ORDER BY up.depth
        """, (file_number, MAX_DEPTH))]

    if not rows:
        return {"chain": [], "ended": "unknown", "detail":
                f"{file_number} is not in this load."}

    _decorate(rows)
    last = rows[-1]
    sup = (last.get("supervisor_id") or "").strip()

    if not sup:
        ended, detail = "top", f"{last['name']} has no supervisor - top of the chain."
    elif len(rows) > MAX_DEPTH:
        ended, detail = "deep", f"Still climbing after {MAX_DEPTH} levels."
    elif any(r["file_number"] == sup for r in rows):
        who = next(r for r in rows if r["file_number"] == sup)
        ended, detail = "cycle", (f"{last['name']} reports to {who['name']} "
                                  f"({sup}), who is already in this chain.")
    else:
        exists = conn.execute(
            "SELECT 1 FROM ADP_Concur_Employees WHERE file_number = ? "
            "AND row_state <> 'deleted'", (sup,)).fetchone()
        if exists:
            # Only reachable if the walk stopped for another reason.
            ended, detail = "deep", f"Chain stopped at {last['name']}."
        else:
            ended, detail = "broken", (
                f"{last['name']} reports to {sup}, who is not in this load"
                + (f" (ADP says \"{last['reports_to_legal_name']}\")"
                   if last.get("reports_to_legal_name") else "") + ".")

    return {"chain": rows, "ended": ended, "detail": detail, "depth": len(rows) - 1}


# ---------------------------------------------------------------- downward


def ADP_Concur_direct_reports(conn: sqlite3.Connection, file_number: str,
                              picked_only: bool = False,
                              status: str = "") -> tuple[list[dict], int]:
    """
    Everyone whose supervisor_id is this person, and how many were filtered out.

    The panel honours the same filter as the tree, so the two cannot disagree
    about how many reports somebody has - but it returns the hidden count as
    well, because 'three reports' when there are really five is the kind of
    quiet lie a filter should never tell.
    """
    rows = _decorate([dict(r) for r in conn.execute(
        f"""
        SELECT {NODE_COLUMNS},
               EXISTS (SELECT 1 FROM ADP_Concur_Selection s
                        WHERE s.employee_key = e.employee_key) AS picked,
               (SELECT COUNT(*) FROM ADP_Concur_Employees c
                 WHERE c.supervisor_id = e.file_number
                   AND c.row_state <> 'deleted') AS n_reports
          FROM ADP_Concur_Employees e
         WHERE e.supervisor_id = ? AND {LIVE}
         ORDER BY e.legal_last_name COLLATE NOCASE, e.legal_first_name COLLATE NOCASE
        """, (file_number,))])
    for r in rows:
        r["picked"] = bool(r["picked"])
    kept = [r for r in rows if wants(r, picked_only, status)]
    return kept, len(rows) - len(kept)


def ADP_Concur_subtree(conn: sqlite3.Connection, file_number: str,
                       include_self: bool = True) -> list[dict]:
    """
    Everyone under one person, at any depth, with the depth on each row.

    This is what 'select this manager's whole organisation' runs, so it has to
    be safe on the real data: the path column stops it revisiting anyone, which
    means a cycle yields each person once instead of forever.
    """
    rows = [dict(r) for r in conn.execute(
        f"""
        WITH RECURSIVE down(employee_key, depth, path) AS (
            SELECT e.employee_key, 0, '/' || e.file_number || '/'
              FROM ADP_Concur_Employees e
             WHERE e.file_number = ? AND {LIVE}
            UNION ALL
            SELECT c.employee_key, down.depth + 1, down.path || c.file_number || '/'
              FROM down
              JOIN ADP_Concur_Employees e ON e.employee_key = down.employee_key
              JOIN ADP_Concur_Employees c ON c.supervisor_id = e.file_number
             WHERE c.row_state <> 'deleted'
               AND down.path NOT LIKE '%/' || c.file_number || '/%'
               AND down.depth < ?
        )
        SELECT {NODE_COLUMNS}, down.depth
          FROM down JOIN ADP_Concur_Employees e ON e.employee_key = down.employee_key
         ORDER BY down.depth,
                  e.legal_last_name COLLATE NOCASE, e.legal_first_name COLLATE NOCASE
        """, (file_number, MAX_DEPTH))]
    _decorate(rows)
    return rows if include_self else [r for r in rows if r["depth"] > 0]


def ADP_Concur_roots(conn: sqlite3.Connection) -> list[dict]:
    """
    Where the tree starts.

    Two kinds of root, and the difference is the whole point: someone with no
    supervisor at all is a real top of the chain, while someone whose
    supervisor is not in the load is a *broken* root - the chain above them was
    meant to continue and does not.
    """
    rows = [dict(r) for r in conn.execute(
        f"""
        SELECT {NODE_COLUMNS},
               CASE WHEN e.supervisor_id = '' THEN 'top' ELSE 'broken' END AS root_kind,
               (SELECT COUNT(*) FROM ADP_Concur_Employees c
                 WHERE c.supervisor_id = e.file_number
                   AND c.row_state <> 'deleted') AS n_reports
          FROM ADP_Concur_Employees e
         WHERE {LIVE}
           AND (e.supervisor_id = ''
                OR NOT EXISTS (SELECT 1 FROM ADP_Concur_Employees s
                                WHERE s.file_number = e.supervisor_id
                                  AND s.row_state <> 'deleted'))
         ORDER BY root_kind,
                  e.legal_last_name COLLATE NOCASE, e.legal_first_name COLLATE NOCASE
        """)]
    return _decorate(rows)


def wants(node: dict, picked_only: bool = False, status: str = "") -> bool:
    """
    Whether one person is somebody the current view is asking to see.

    Every tree filter goes through here, so they combine rather than compete:
    'active only' and 'only picked' together means the picked people who are
    also active, and adding a third filter later is one more clause here.
    """
    if picked_only and not node.get("picked"):
        return False
    if status == "active" and node.get("concur_status") != "Y":
        return False
    if status == "inactive" and node.get("concur_status") != "N":
        return False
    return True


def ADP_Concur_forest(conn: sqlite3.Connection, picked_only: bool = False,
                      status: str = "") -> list[dict]:
    """
    The whole hierarchy as nested nodes, roots first.

    Built in Python from two flat queries rather than one recursive walk,
    because the tree has to include people a recursive walk would never reach -
    anyone inside a cycle has no root to descend from. Those are attached at
    the end under their own heading rather than quietly dropped.

    Filtering prunes the tree but never breaks it: a manager who fails the
    filter is kept when somebody below them passes it, because a tree with the
    middle cut out of it is not a tree. A leaver who still has active people
    under him is exactly that case, and hiding him would orphan his whole
    branch. Every node carries `wanted` - true if it passed the filter itself,
    false if it is only there to hold a branch up - so the page can show the
    difference rather than implying everything on screen matched.
    """
    rows = {r["file_number"]: dict(r) for r in conn.execute(
        f"""
        SELECT {NODE_COLUMNS},
               EXISTS (SELECT 1 FROM ADP_Concur_Selection s
                        WHERE s.employee_key = e.employee_key) AS picked
          FROM ADP_Concur_Employees e WHERE {LIVE}
         ORDER BY e.legal_last_name COLLATE NOCASE, e.legal_first_name COLLATE NOCASE
        """)}
    for node in rows.values():
        node["name"] = _name(node)
        node["picked"] = bool(node["picked"])
        node["wanted"] = wants(node, picked_only, status)
        node["children"] = []

    if picked_only or status:
        # Keep everyone the filter wants, and every ancestor of one. Walking up
        # from each match is cheaper than walking the whole tree down, and
        # easier to be sure is right: a node survives exactly when it matched
        # or holds somebody who did.
        keep: set[str] = set()
        for fn, node in rows.items():
            if not node["wanted"]:
                continue
            cur, seen = fn, set()
            while cur and cur not in seen:
                seen.add(cur)
                keep.add(cur)
                cur = (rows[cur].get("supervisor_id") or "").strip() \
                    if cur in rows else ""
        rows = {fn: n for fn, n in rows.items() if fn in keep}

    roots = []
    for node in rows.values():
        sup = (node.get("supervisor_id") or "").strip()
        parent = rows.get(sup) if sup else None
        if parent is not None and parent is not node:
            parent["children"].append(node)
        else:
            node["root_kind"] = "top" if not sup else "broken"
            roots.append(node)

    # Anyone not reachable from a root is inside a loop. Surface them.
    seen: set[str] = set()
    stack = list(roots)
    while stack:
        node = stack.pop()
        if node["file_number"] in seen:
            continue
        seen.add(node["file_number"])
        stack.extend(node["children"])

    for fn, node in rows.items():
        if fn not in seen:
            node["root_kind"] = "cycle"
            roots.append(node)

    def counted(node: dict) -> int:
        node["n_subtree"] = 1 + sum(counted(c) for c in node["children"])
        node["n_reports"] = len(node["children"])
        return node["n_subtree"]

    for r in roots:
        if r.get("root_kind") != "cycle":
            counted(r)
        else:
            r["n_subtree"] = 1
            r["n_reports"] = len(r["children"])

    roots.sort(key=lambda r: ({"top": 0, "broken": 1, "cycle": 2}[r["root_kind"]],
                              -r["n_subtree"], r["name"]))
    return roots


# ---------------------------------------------------------------- problems


def ADP_Concur_hierarchy_problems(conn: sqlite3.Connection) -> dict:
    """
    Everything wrong with the chain, in the three shapes it goes wrong in.

    Called by the derive so these land in the exception list beside the
    mapping failures, and by the Hierarchy tab so they can be looked at as a
    group. Nobody appears in more than one list.
    """
    self_led = _decorate([dict(r) for r in conn.execute(
        f"SELECT {NODE_COLUMNS} FROM ADP_Concur_Employees e "
        f"WHERE {LIVE} AND e.supervisor_id <> '' "
        f"AND e.supervisor_id = e.file_number")])
    self_keys = {r["employee_key"] for r in self_led}

    broken = [r for r in _decorate([dict(r) for r in conn.execute(
        f"""
        SELECT {NODE_COLUMNS} FROM ADP_Concur_Employees e
         WHERE {LIVE} AND e.supervisor_id <> ''
           AND NOT EXISTS (SELECT 1 FROM ADP_Concur_Employees s
                            WHERE s.file_number = e.supervisor_id
                              AND s.row_state <> 'deleted')
         ORDER BY e.legal_last_name COLLATE NOCASE
        """)]) if r["employee_key"] not in self_keys]

    # A cycle is anyone the forest could not reach from a root.
    reachable: set[str] = set()
    stack = list(ADP_Concur_forest(conn))
    while stack:
        node = stack.pop()
        if node["file_number"] in reachable:
            continue
        reachable.add(node["file_number"])
        stack.extend(node["children"])
    cycles = _decorate([dict(r) for r in conn.execute(
        f"SELECT {NODE_COLUMNS} FROM ADP_Concur_Employees e WHERE {LIVE}")])
    cycles = [r for r in cycles
              if r["file_number"] not in reachable and r["employee_key"] not in self_keys]

    return {"broken": broken, "self_led": self_led, "cycles": cycles,
            "totals": {"broken": len(broken), "self_led": len(self_led),
                       "cycles": len(cycles)}}


def ADP_Concur_validate_supervisor_map(conn: sqlite3.Connection) -> dict:
    """
    Check the Supervisor Map against the employees, row by row.

    The map is an override the derive always honours, which makes it powerful
    and makes a wrong row expensive: it is the one place where a typo silently
    rewrites the approval routing for somebody. Two things can be wrong with a
    row and neither shows up by reading it -

      the supervisor it names is not an employee, so the row hands Concur an
      approver that does not exist; or
      the employee it is keyed on is not in the load, so the row is dead
      weight that will never fire, and whoever wrote it thinks it did.

    A verdict per row:

      ok                  both ends resolve
      top                 supervisor deliberately blank - top of the chain
      supervisor_missing  the supervisor is not in the load
      employee_missing    the employee is not in the load; the row never applies
      self                keyed on the same person it names as supervisor
      name_mismatch       both ends resolve but a name in the row disagrees
                          with the employee it points at - usually a row keyed
                          on the wrong File Number
    """
    rows = [dict(r) for r in conn.execute(
        """
        SELECT m.*,
               e.employee_key AS emp_key,
               e.legal_last_name || ', ' || e.legal_first_name AS emp_actual,
               e.business_unit_desc, e.business_unit_code, e.home_department_code,
               e.home_department_desc, e.location_desc, e.legal_country_code,
               e.language_desc, e.pay_grade_code, e.payroll_company_code,
               e.employee_type, e.pay_frequency,
               s.employee_key AS sup_key,
               s.legal_last_name || ', ' || s.legal_first_name AS sup_actual
          FROM ADP_Concur_SupervisorMap m
          LEFT JOIN ADP_Concur_Employees e
                 ON e.file_number = m.file_number AND e.row_state <> 'deleted'
          LEFT JOIN ADP_Concur_Employees s
                 ON s.file_number = m.supervisor_id AND s.row_state <> 'deleted'
         ORDER BY m.file_number
        """)]

    def matches(written: str, actual: str) -> bool:
        """
        Loose enough not to cry over 'Bailey, Jason' vs 'Bailey,Jason'.

        Compares the set of alphabetic runs, so punctuation, spacing and order
        are all forgiven and only a genuinely different name is reported.
        """
        norm = lambda v: {p for p in re.split(r"[^A-Za-z]+", (v or "").lower()) if p}
        w, a = norm(written), norm(actual)
        return not w or not a or bool(w & a)

    out = []
    for r in rows:
        sup_id = (r["supervisor_id"] or "").strip()
        verdict, detail = "ok", ""

        if r["emp_key"] is None:
            verdict = "employee_missing"
            detail = (f"{r['file_number']} is not in this load, so this row never "
                      "applies to anybody.")
        elif not sup_id:
            verdict = "top"
            detail = (f"{r['emp_actual']} is deliberately left with no supervisor"
                      + (f" — \"{r['note']}\"." if r["note"] else "."))
        elif sup_id == (r["file_number"] or "").strip():
            verdict = "self"
            detail = f"{r['emp_actual']} is mapped to themselves."
        elif r["sup_key"] is None:
            verdict = "supervisor_missing"
            detail = (f"Supervisor {sup_id} is not in this load"
                      + (f" — the map calls them \"{r['supervisor_name']}\""
                         if r["supervisor_name"] else "")
                      + (". That is not a File Number, so it looks like a "
                         "placeholder." if not sup_id.isdigit() else "."))
        elif not matches(r["employee_name"], r["emp_actual"]):
            verdict = "name_mismatch"
            detail = (f"The row says \"{r['employee_name']}\" but {r['file_number']} "
                      f"is {r['emp_actual']} — the row may be keyed on the wrong "
                      "File Number.")
        elif not matches(r["supervisor_name"], r["sup_actual"]):
            verdict = "name_mismatch"
            detail = (f"The row calls {sup_id} \"{r['supervisor_name']}\" but "
                      f"{sup_id} is {r['sup_actual']}.")

        r["verdict"] = verdict
        r["detail"] = detail
        r["placeholder"] = bool(sup_id) and not sup_id.isdigit()
        r["severity"] = ("error" if verdict in ("supervisor_missing", "self")
                         else "warning" if verdict in ("employee_missing",
                                                       "name_mismatch")
                         else "ok")
        out.append(r)

    totals: dict = {"rows": len(out)}
    for r in out:
        totals[r["verdict"]] = totals.get(r["verdict"], 0) + 1
    totals["problems"] = sum(1 for r in out if r["severity"] != "ok")
    return {"rows": out, "totals": totals}


def ADP_Concur_hierarchy_stats(conn: sqlite3.Connection) -> dict:
    """Shape of the tree: how many roots, how deep, who has the most people."""
    forest = ADP_Concur_forest(conn)
    depths: dict[int, int] = {}

    def walk(node, depth):
        depths[depth] = depths.get(depth, 0) + 1
        for c in node["children"]:
            walk(c, depth + 1)

    for r in forest:
        walk(r, 0)

    managers = [dict(r) for r in conn.execute(
        """
        SELECT e.file_number,
               e.legal_last_name || ', ' || e.legal_first_name AS name,
               e.job_title,
               (SELECT COUNT(*) FROM ADP_Concur_Employees c
                 WHERE c.supervisor_id = e.file_number
                   AND c.row_state <> 'deleted') AS n_reports
          FROM ADP_Concur_Employees e
         WHERE e.row_state <> 'deleted'
         ORDER BY n_reports DESC, name LIMIT 12
        """)]
    return {"roots": len(forest),
            "tops": sum(1 for r in forest if r.get("root_kind") == "top"),
            "broken_roots": sum(1 for r in forest if r.get("root_kind") == "broken"),
            "max_depth": max(depths) if depths else 0,
            "by_depth": [{"depth": d, "n": depths[d]} for d in sorted(depths)],
            "managers": [m for m in managers if m["n_reports"]]}


def ADP_Concur_subtree_keys(conn: sqlite3.Connection,
                            file_numbers: list[str]) -> list[int]:
    """
    Employee keys for everyone under any of these people, plus the people
    themselves. What 'select this whole organisation' resolves to before the
    extract is scoped by it.
    """
    keys: list[int] = []
    seen: set[int] = set()
    for fn in file_numbers:
        for row in ADP_Concur_subtree(conn, fn):
            if row["employee_key"] not in seen:
                seen.add(row["employee_key"])
                keys.append(row["employee_key"])
    return keys


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Walk the Concur supervisor chain.")
    ap.add_argument("--db", default=None, help=f"SQLite path (default {DEFAULT_DB_PATH})")
    ap.add_argument("--chain", metavar="FILE_NUMBER",
                    help="Print one employee's chain up to the top")
    ap.add_argument("--subtree", metavar="FILE_NUMBER",
                    help="Print everyone under one employee")
    ap.add_argument("--problems", action="store_true",
                    help="List broken links, self-supervision and loops")
    ap.add_argument("--stats", action="store_true", help="Shape of the tree")
    ap.add_argument("--map", action="store_true",
                    help="Validate the Supervisor Map against the employees")
    args = ap.parse_args(argv)

    conn = connect(args.db)
    print(f"Database: {resolve_db_path(args.db)}")

    if args.chain:
        res = ADP_Concur_chain_up(conn, args.chain)
        for r in res["chain"]:
            print(f"  {'  ' * r['depth']}{'└ ' if r['depth'] else ''}"
                  f"{r['file_number']}  {r['name']}  ({r['job_title'] or ''})")
        print(f"  ends: {res['ended']} - {res['detail']}")

    if args.subtree:
        rows = ADP_Concur_subtree(conn, args.subtree)
        for r in rows:
            print(f"  {'  ' * r['depth']}{r['file_number']}  {r['name']}")
        print(f"  {len(rows)} people including the top")

    if args.problems:
        p = ADP_Concur_hierarchy_problems(conn)
        for kind, label in (("broken", "supervisor not in this load"),
                            ("self_led", "reports to themselves"),
                            ("cycles", "inside a loop")):
            print(f"\n{p['totals'][kind]} {label}")
            for r in p[kind]:
                print(f"  {r['file_number']}  {r['name']:<28} -> {r['supervisor_id']}"
                      f"  (ADP: {r['reports_to_legal_name'] or '-'})")

    if args.map:
        v = ADP_Concur_validate_supervisor_map(conn)
        t = v["totals"]
        print(f"\n{t['rows']} Supervisor Map row(s), {t['problems']} with a problem")
        for r in v["rows"]:
            mark = {"error": "!!", "warning": " ?", "ok": "  "}[r["severity"]]
            print(f"  {mark} {r['file_number']:<8} {r['verdict']:<19} "
                  f"{r['detail'] or 'both ends resolve'}")

    if args.stats:
        s = ADP_Concur_hierarchy_stats(conn)
        print(f"\n{s['roots']} root(s): {s['tops']} genuine top(s), "
              f"{s['broken_roots']} broken; {s['max_depth']} level(s) deep")
        for d in s["by_depth"]:
            print(f"  level {d['depth']}: {d['n']}")
        print("\nlargest spans of control")
        for m in s["managers"]:
            print(f"  {m['n_reports']:>4}  {m['name']}  ({m['job_title'] or ''})")

    conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
