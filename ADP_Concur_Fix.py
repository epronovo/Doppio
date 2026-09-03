"""
ADP_Concur_Fix - turn the exception list into a list of things you can do.

The exceptions say what is wrong one employee at a time. That is the wrong
grain for actually clearing them: two people blocked by the same absent
supervisor are one missing person, not two problems, and eleven people in an
unmapped department are one missing Org Map row. This module regroups them by
*what would fix them*, works out what the fix should be filled in with, and
applies it.

Every remedy is expressed as data - an action name and a payload of fields -
so the front end renders a form from it rather than knowing anything about
supervisors or department codes, and the same fixes are available from the
command line.

Nothing here writes a derived column. Fixes write to the things the derive
reads from: the employee's own ADP fields, the six map tables, and the config.
Then the derive runs and the exception list is rebuilt, which is how a fix
proves it worked - the count goes down.
"""
from __future__ import annotations

import argparse
import sqlite3
from collections import Counter

from ADP_Concur_Db import DEFAULT_DB_PATH, connect, resolve_db_path
from ADP_Concur_Map import (
    ADP_Concur_derive,
    UNMAPPED,
    load_config,
    save_config,
)

# Which map holds the answer for each unmapped value, what the key column is,
# and which employee field the unmapped key came from. This table is the whole
# of the 'add the missing map row' logic.
UNMAPPED_FIXES = {
    "org_unit_1": {
        "map": "org", "table": "ADP_Concur_OrgMap",
        "key_column": "business_unit_desc", "employee_field": "business_unit_desc",
        "label": "Business Unit Description",
        "fields": ["business_unit_desc", "home_department_desc", "home_department_code",
                   "org_unit_1", "org_unit_2", "default_language", "currency"],
    },
    "org_unit_2": {
        "map": "org", "table": "ADP_Concur_OrgMap",
        "key_column": "home_department_code", "employee_field": "home_department_code",
        "label": "Home Department Code",
        "fields": ["business_unit_desc", "home_department_desc", "home_department_code",
                   "org_unit_1", "org_unit_2", "default_language", "currency"],
    },
    "reimbursement_currency": {
        "map": "org", "table": "ADP_Concur_OrgMap",
        "key_column": "business_unit_desc", "employee_field": "business_unit_desc",
        "label": "Business Unit Description",
        "fields": ["business_unit_desc", "home_department_desc", "home_department_code",
                   "org_unit_1", "org_unit_2", "default_language", "currency"],
    },
    "concur_profile": {
        "map": "salary", "table": "ADP_Concur_SalaryMap",
        "key_column": "pay_grade_code", "employee_field": "pay_grade_code",
        "label": "Pay Grade Code",
        "fields": ["pay_grade_code", "pay_grade_desc", "expense_map", "travel_map"],
    },
    "travel_profile": {
        "map": "salary", "table": "ADP_Concur_SalaryMap",
        "key_column": "pay_grade_code", "employee_field": "pay_grade_code",
        "label": "Pay Grade Code",
        "fields": ["pay_grade_code", "pay_grade_desc", "expense_map", "travel_map"],
    },
    "legal_country": {
        "map": "country", "table": "ADP_Concur_CountryMap",
        "key_column": "adp_country", "employee_field": "legal_country_code",
        "label": "Country Code",
        "fields": ["adp_country", "concur_country"],
    },
    "locale_code": {
        "map": "language", "table": "ADP_Concur_LanguageMap",
        "key_column": "language_desc", "employee_field": "language_desc",
        "label": "Language Description",
        "fields": ["language_desc", "language_code", "adp_language"],
    },
    "concur_status": {
        "map": "status", "table": "ADP_Concur_StatusMap",
        "key_column": "position_status", "employee_field": "position_status",
        "label": "Position Status",
        "fields": ["position_status", "concur_status"],
    },
}

# The map tables a fix may write to, and their key columns - so a fix can only
# reach a table this module means it to.
MAP_TABLES = {
    "org": ("ADP_Concur_OrgMap",
            ["business_unit_desc", "home_department_desc", "home_department_code",
             "org_unit_1", "org_unit_2", "default_language", "currency"]),
    "status": ("ADP_Concur_StatusMap", ["position_status", "concur_status"]),
    "country": ("ADP_Concur_CountryMap", ["adp_country", "concur_country"]),
    "language": ("ADP_Concur_LanguageMap",
                 ["language_desc", "language_code", "adp_language"]),
    "salary": ("ADP_Concur_SalaryMap",
               ["pay_grade_code", "pay_grade_desc", "expense_map", "travel_map"]),
    "supervisor": ("ADP_Concur_SupervisorMap",
                   ["file_number", "employee_name", "supervisor_name",
                    "supervisor_id", "note"]),
}


def _s(v) -> str:
    return "" if v is None else str(v).strip()


def _common(values: list[str]) -> str:
    """The most frequent non-blank value, for filling a guess from the group."""
    vals = [v for v in values if _s(v)]
    return Counter(vals).most_common(1)[0][0] if vals else ""


def _split_name(legal_name: str) -> tuple[str, str]:
    """
    'Erdogan, Mehmet' -> ('Erdogan', 'Mehmet'); 'Mehmet Erdogan' -> the same.

    ADP writes Reports To Legal Name last-first with a comma, which is the
    only reason a missing supervisor can be created with a real name on it.
    """
    name = _s(legal_name)
    if not name:
        return "", ""
    if "," in name:
        last, _, first = name.partition(",")
        return last.strip(), first.strip()
    parts = name.split()
    return (parts[-1], " ".join(parts[:-1])) if len(parts) > 1 else (name, "")


# ------------------------------------------------------------------ the plan


def _missing_people(conn: sqlite3.Connection) -> list[dict]:
    """
    Supervisor IDs that at least one employee points at and nobody answers to.

    Grouped by the ID, because that is the unit of work: 207199 is one person
    to create, however many people report to them.
    """
    rows = [dict(r) for r in conn.execute(
        """
        SELECT e.employee_key, e.file_number, e.supervisor_id,
               e.legal_last_name || ', ' || e.legal_first_name AS name,
               e.reports_to_legal_name, e.business_unit_desc, e.business_unit_code,
               e.home_department_code, e.home_department_desc, e.location_desc,
               e.legal_country_code, e.language_desc, e.pay_grade_code,
               e.payroll_company_code, e.employee_type, e.pay_frequency
          FROM ADP_Concur_Employees e
         WHERE e.row_state <> 'deleted' AND e.supervisor_id <> ''
           AND NOT EXISTS (SELECT 1 FROM ADP_Concur_Employees s
                            WHERE s.file_number = e.supervisor_id
                              AND s.row_state <> 'deleted')
         ORDER BY e.supervisor_id, e.legal_last_name
        """)]

    groups: dict[str, list[dict]] = {}
    for r in rows:
        groups.setdefault(r["supervisor_id"], []).append(r)

    # Where the Supervisor Map is what put an ID on somebody, the map's own
    # supervisor_name is the better source for the name - it was typed by
    # whoever knew who they meant, and ADP's Reports To may be stale or blank.
    map_names = {r["supervisor_id"]: r["supervisor_name"] for r in conn.execute(
        "SELECT supervisor_id, supervisor_name FROM ADP_Concur_SupervisorMap "
        "WHERE supervisor_id <> '' AND supervisor_name <> ''")}

    out = []
    for sup_id, reporters in groups.items():
        named = map_names.get(sup_id) or _common(
            [r["reports_to_legal_name"] for r in reporters])
        last, first = _split_name(named)
        # A manager almost always shares a business unit with their reports, so
        # the group is a better source for the new record than nothing at all.
        # Every guessed field is listed in `guessed` and shown as such.
        prefill = {
            "file_number": sup_id,
            "legal_last_name": last,
            "legal_first_name": first,
            "reports_to_legal_name": "",
            "position_status": "Active",
            "payroll_company_code": _common([r["payroll_company_code"] for r in reporters]),
            "business_unit_code": _common([r["business_unit_code"] for r in reporters]),
            "business_unit_desc": _common([r["business_unit_desc"] for r in reporters]),
            "home_department_code": _common([r["home_department_code"] for r in reporters]),
            "home_department_desc": _common([r["home_department_desc"] for r in reporters]),
            "location_desc": _common([r["location_desc"] for r in reporters]),
            "legal_country_code": _common([r["legal_country_code"] for r in reporters]),
            "language_desc": _common([r["language_desc"] for r in reporters]),
            "employee_type": _common([r["employee_type"] for r in reporters]),
            "pay_frequency": _common([r["pay_frequency"] for r in reporters]),
            # A manager's grade is at least their reports', so the highest in
            # the group is the least wrong starting point.
            "pay_grade_code": max([_s(r["pay_grade_code"]) for r in reporters] or [""]),
            "work_email": "",
            "job_title": "",
        }
        looks_like_id = sup_id.isdigit()
        out.append({
            "kind": "missing_person",
            "key": sup_id,
            "severity": "error",
            "title": f"{sup_id}" + (f" — {named}" if named else ""),
            "detail":
                (f"{_people(len(reporters))} "
                 + ("reports" if len(reporters) == 1 else "report")
                 + f" to {sup_id}, who is not in this load."
                 + ("" if looks_like_id else
                    f" — and {sup_id} is not a File Number at all, so it is "
                    "almost certainly a placeholder somebody typed into the "
                    "Supervisor Map.")),
            "n_blocked": len(reporters),
            "employees": [{"employee_key": r["employee_key"],
                           "file_number": r["file_number"], "name": r["name"]}
                          for r in reporters],
            "remedies": ([{
                "action": "create_employee",
                "label": f"Create {named or sup_id}",
                "note": "Added by hand, exactly like anyone else ADP does not "
                        "have yet. The greyed values are guesses taken from the "
                        "people who report to them - check them.",
                "fields": prefill,
                "guessed": ["payroll_company_code", "business_unit_code",
                            "business_unit_desc", "home_department_code",
                            "home_department_desc", "location_desc",
                            "legal_country_code", "language_desc",
                            "employee_type", "pay_frequency", "pay_grade_code"],
            }] if looks_like_id else []) + [
                {
                    "action": "set_supervisor",
                    "label": "Point them at somebody already in the load",
                    "note": "Writes a Supervisor Map row for each of them, which "
                            "overrides whatever ADP sent and survives the next "
                            "import.",
                    "fields": {"supervisor_id": "",
                               "file_numbers": [r["file_number"] for r in reporters]},
                },
                {
                    "action": "set_supervisor",
                    "label": "Leave them with no supervisor",
                    "note": "A blank Supervisor Map row is how 'top of the food "
                            "chain' is expressed - the same thing Teresa Bair has.",
                    "fields": {"supervisor_id": "",
                               "file_numbers": [r["file_number"] for r in reporters],
                               "note": "Top of food chain"},
                    "confirm": True,
                },
            ],
        })
    out.sort(key=lambda g: -g["n_blocked"])
    return out


def _people(n: int) -> str:
    return "1 person" if n == 1 else f"{n} people"


def _have(n: int) -> str:
    return "has" if n == 1 else "have"


def _incomplete_group(spec: dict, field: str, people: list[dict]) -> dict:
    """
    ADP left the field blank.

    Merged across fields by the front end where the same person appears twice -
    an employee with no business unit usually has no pay grade either, and
    Carla Cunningham in this workbook has neither. Two remedies, and which one
    is right depends on why the record is empty: fill it in, or accept that a
    leaver nobody is going to reimburse does not need a Concur profile.
    """
    field_label = spec["employee_field"]
    return {
        "kind": "incomplete",
        "employee_field": spec["employee_field"],
        "map": spec["map"],
        "key": "",
        "field": field,
        "fields_affected": [field],
        "severity": "error",
        "title": f"{_people(len(people))} {_have(len(people))} no "
                 f"{spec['label']} in ADP",
        "detail":
            f"The {spec['label']} column is empty on their ADP record, so "
            f"nothing can be looked up for them. Either the record needs "
            f"finishing, or it is one of the leavers that can be left out.",
        "n_blocked": len(people),
        "employees": [{"employee_key": p["employee_key"],
                       "file_number": p["file_number"], "name": p["name"]}
                      for p in people],
        "remedies": [
            {
                "action": "set_employee_fields",
                "label": f"Fill in the {spec['label']}",
                "note": "Written onto the employee here. The next ADP import "
                        "overwrites it with whatever ADP says by then, so this "
                        "is a stopgap unless it is also fixed at source.",
                "per_employee": True,
                "fields": {"file_number": "", field_label: ""},
                "employees": [{"file_number": p["file_number"], "name": p["name"]}
                              for p in people],
            },
            {
                "action": "exclude_employee",
                "label": "Leave them out of the load",
                "note": "The record stays in the database and simply is not "
                        "written to the file, so it can be put back.",
                "per_employee": True,
                "fields": {"file_number": ""},
                "employees": [{"file_number": p["file_number"], "name": p["name"]}
                              for p in people],
                "confirm": True,
            },
        ],
    }


def _unmapped_groups(conn: sqlite3.Connection) -> list[dict]:
    """
    Every '... Not Mapped' value, grouped by the value that is missing.

    The key is read off the employee rather than parsed out of the message, so
    a reworded message never breaks the fix.
    """
    out = []
    for field, spec in UNMAPPED_FIXES.items():
        rows = [dict(r) for r in conn.execute(
            f"""
            SELECT e.employee_key, e.file_number,
                   e.legal_last_name || ', ' || e.legal_first_name AS name,
                   e.{spec['employee_field']} AS key_value,
                   e.business_unit_desc, e.home_department_desc,
                   e.home_department_code, e.pay_grade_desc, e.concur_status
              FROM ADP_Concur_Employees e
             WHERE e.row_state <> 'deleted' AND e.{field} LIKE ?
             ORDER BY e.legal_last_name
            """, (f"%{UNMAPPED[field]}%",))]
        groups: dict[str, list[dict]] = {}
        for r in rows:
            groups.setdefault(_s(r["key_value"]), []).append(r)

        for key, people in groups.items():
            # A blank key is not a missing map row - it is a missing value on
            # the employee, and no Org Map row keyed on the empty string would
            # ever be right. Kelly's "bad or incomplete data in ADP" lands
            # here, and the remedy is to fill it in or leave the person out.
            if not key:
                # Two derived fields can share one employee column - a blank
                # Business Unit Description breaks org_unit_1 and the currency
                # alike - and that is one empty cell, not two problems.
                if any(g["kind"] == "incomplete"
                       and g["employee_field"] == spec["employee_field"]
                       for g in out):
                    continue
                out.append(_incomplete_group(spec, field, people))
                continue

            # One row per missing key, not per field: a business unit missing
            # from the Org Map breaks org_unit_1 and the currency together, and
            # adding the row fixes both.
            existing = next((g for g in out if g["kind"] == "unmapped"
                             and g["map"] == spec["map"] and g["key"] == key), None)
            if existing:
                existing["fields_affected"].append(field)
                continue

            prefill = {c: "" for c in spec["fields"]}
            prefill[spec["key_column"]] = key
            # Fill the columns the employees themselves can answer for, and
            # borrow the rest of an Org Map row from the same business unit.
            if spec["map"] == "org":
                bu = _common([p["business_unit_desc"] for p in people])
                prefill["business_unit_desc"] = prefill["business_unit_desc"] or bu
                prefill["home_department_code"] = \
                    prefill["home_department_code"] or _common(
                        [p["home_department_code"] for p in people])
                prefill["home_department_desc"] = \
                    prefill["home_department_desc"] or _common(
                        [p["home_department_desc"] for p in people])
                sibling = conn.execute(
                    "SELECT org_unit_1, default_language, currency "
                    "FROM ADP_Concur_OrgMap WHERE business_unit_desc = ? LIMIT 1",
                    (bu,)).fetchone()
                if sibling:
                    prefill["org_unit_1"] = sibling["org_unit_1"]
                    prefill["default_language"] = sibling["default_language"]
                    prefill["currency"] = sibling["currency"]
            elif spec["map"] == "salary":
                prefill["pay_grade_desc"] = _common([p["pay_grade_desc"] for p in people])
            elif key and spec["map"] == "country" and len(key) >= 2:
                prefill["concur_country"] = key[:2].upper()

            blank = [c for c in spec["fields"] if not prefill[c]]
            out.append({
                "kind": "unmapped",
                "map": spec["map"],
                "key": key,
                "field": field,
                "fields_affected": [field],
                "severity": "error",
                "title": f"{spec['label']} “{key or '(blank)'}” is not mapped",
                "detail":
                    f"{_people(len(people))} carry this value and the "
                    f"{spec['map'].title()} Map has no row for it, so every "
                    "value it feeds comes out unmapped.",
                "n_blocked": len(people),
                "employees": [{"employee_key": p["employee_key"],
                               "file_number": p["file_number"], "name": p["name"]}
                              for p in people],
                "remedies": [{
                    "action": "add_map_row",
                    "label": f"Add the {spec['map'].title()} Map row",
                    "note": ("Everything blank needs an answer: "
                             + ", ".join(c.replace("_", " ") for c in blank)
                             if blank else "Check the values and save."),
                    "map": spec["map"],
                    "fields": prefill,
                    "required": blank,
                }],
            })
    out.sort(key=lambda g: -g["n_blocked"])
    return out


def _login_group(conn: sqlite3.Connection) -> list[dict]:
    """
    People with no Login ID, as one decision rather than thirty-four.

    The remedy that actually scales is the rule, not the records: pick a
    fallback and every one of them gets an ID. Editing the people one at a
    time is offered too, because for a handful a real address is better.
    """
    people = [dict(r) for r in conn.execute(
        """
        SELECT e.employee_key, e.file_number,
               e.legal_last_name || ', ' || e.legal_first_name AS name,
               e.personal_email
          FROM ADP_Concur_Employees e
         WHERE e.row_state <> 'deleted' AND (e.login_id IS NULL OR e.login_id = '')
         ORDER BY e.legal_last_name
        """)]
    if not people:
        return []

    with_personal = sum(1 for p in people if _s(p["personal_email"]))
    cfg = load_config()
    sources = (cfg.get("login_id") or {}).get("sources") or ["work_email"]
    return [{
        "kind": "login_id",
        "key": "login_id",
        "severity": "error",
        "title": f"{_people(len(people))} {_have(len(people))} no Login ID",
        "detail":
            "ADP has no work email address for them"
            + (f", though all {with_personal} have a personal one"
               if with_personal == len(people)
               else f"; {with_personal} of them do have a personal one")
            + ". Nothing is invented on their behalf, so they are held out of "
              "the extract until the rule says what to fall back to.",
        "n_blocked": len(people),
        "employees": [{"employee_key": p["employee_key"],
                       "file_number": p["file_number"], "name": p["name"]}
                      for p in people],
        "remedies": [{
            "action": "set_login_sources",
            "label": "Change the Login ID rule for everyone",
            "note": "The order the rule tries. Adding personal_email covers "
                    f"{with_personal} of them; file_number covers the rest, and "
                    "needs a domain to hang the bare number on.",
            "fields": {"sources": ", ".join(sources),
                       "bare_domain": (cfg.get("login_id") or {}).get("bare_domain", "")},
        }],
    }]


def _supervisor_map_groups(conn: sqlite3.Connection) -> list[dict]:
    """
    Rows in the Supervisor Map that do not hold up.

    The employee-side checks cannot see these. A row keyed on somebody who is
    not in the load never fires, so nothing anywhere reports it - and whoever
    wrote it believes it is doing something. A row whose names disagree with
    the employees at either end usually means it was keyed on the wrong File
    Number, which is worse than useless: it is silently rewriting the approval
    routing for the wrong person.

    The map's supervisor-missing rows are deliberately *not* repeated here -
    they already surface as a missing person, with the create form on them.
    """
    from ADP_Concur_Hierarchy import ADP_Concur_validate_supervisor_map

    rows = ADP_Concur_validate_supervisor_map(conn)["rows"]
    out = []

    dead = [r for r in rows if r["verdict"] == "employee_missing"]
    if dead:
        out.append({
            "kind": "supervisor_map_dead", "key": "map_dead", "severity": "warning",
            "title": f"{_people(len(dead))} in the Supervisor Map "
                     f"{_have(len(dead))} no employee record",
            "detail": "These rows are keyed on a File Number this load does not "
                      "have, so they never apply to anybody. Either the person "
                      "belongs in the load, or the row is left over.",
            "n_blocked": len(dead),
            "employees": [{"employee_key": None, "file_number": r["file_number"],
                           "name": r["employee_name"] or "(not in the load)"}
                          for r in dead],
            "remedies": [
                {
                    "action": "create_employee",
                    "label": "Create the employee the row is about",
                    "note": "The map already names them; everything else has to "
                            "be filled in, because nothing in this load knows "
                            "anything more about them.",
                    "per_employee": True,
                    "fields": {"file_number": "", "legal_last_name": "",
                               "legal_first_name": "", "position_status": "Active"},
                    "employees": [{"file_number": r["file_number"],
                                   "name": r["employee_name"] or r["file_number"]}
                                  for r in dead],
                },
                {
                    "action": "delete_map_row",
                    "label": "Remove the row",
                    "note": "It is not doing anything. Removing it keeps the map "
                            "honest about what it is actually overriding.",
                    "per_employee": True,
                    "map": "supervisor",
                    "fields": {"file_number": ""},
                    "employees": [{"file_number": r["file_number"],
                                   "name": r["employee_name"] or r["file_number"]}
                                  for r in dead],
                    "confirm": True,
                },
            ],
        })

    selfies = [r for r in rows if r["verdict"] == "self"]
    if selfies:
        out.append({
            "kind": "supervisor_map_self", "key": "map_self", "severity": "error",
            "title": f"{_people(len(selfies))} in the Supervisor Map "
                     f"{_have(len(selfies))} themselves as supervisor",
            "detail": "Concur will not accept a record as its own approver.",
            "n_blocked": len(selfies),
            "employees": [{"employee_key": r["emp_key"],
                           "file_number": r["file_number"],
                           "name": r["emp_actual"] or r["employee_name"]}
                          for r in selfies],
            "remedies": [{
                "action": "set_supervisor",
                "label": "Point them at somebody else",
                "note": "Rewrites the same Supervisor Map row.",
                "fields": {"supervisor_id": "",
                           "file_numbers": [r["file_number"] for r in selfies]},
            }],
        })

    mismatched = [r for r in rows if r["verdict"] == "name_mismatch"]
    if mismatched:
        out.append({
            "kind": "supervisor_map_names", "key": "map_names", "severity": "warning",
            "title": f"{_people(len(mismatched))} in the Supervisor Map "
                     f"{_have(len(mismatched))} a name that does not match",
            "detail": "Both ends resolve, so the routing works — but a name in "
                      "the row disagrees with the employee it points at, which "
                      "usually means the row was keyed on the wrong File Number "
                      "and is quietly rewriting the wrong person's approver. "
                      "Worth reading before trusting it.",
            "n_blocked": len(mismatched),
            "employees": [{"employee_key": r["emp_key"],
                           "file_number": r["file_number"],
                           "name": r["detail"]} for r in mismatched],
            "remedies": [
                {
                    "action": "set_supervisor",
                    "label": "Set the supervisor deliberately",
                    "note": "Rewrites the row, names and all, from what is "
                            "actually in the load.",
                    "fields": {"supervisor_id": "",
                               "file_numbers": [r["file_number"] for r in mismatched]},
                },
                {
                    "action": "delete_map_row",
                    "label": "Remove the row",
                    "note": "Falls back to whatever ADP sent for these people.",
                    "per_employee": True,
                    "map": "supervisor",
                    "fields": {"file_number": ""},
                    "employees": [{"file_number": r["file_number"],
                                   "name": r["emp_actual"] or r["file_number"]}
                                  for r in mismatched],
                    "confirm": True,
                },
            ],
        })

    return out


def _chain_groups(conn: sqlite3.Connection) -> list[dict]:
    """Self-supervision and loops - rarer, same remedy shape as the rest."""
    from ADP_Concur_Hierarchy import ADP_Concur_hierarchy_problems

    problems = ADP_Concur_hierarchy_problems(conn)
    out = []
    for kind, rows, title, detail in (
        ("self_supervisor", problems["self_led"], "reports to themselves",
         "Concur will not accept a record as its own approver."),
        ("cycle", problems["cycles"], "sit inside a supervisor loop",
         "Following the chain upwards comes back to where it started rather "
         "than reaching the top."),
    ):
        if not rows:
            continue
        out.append({
            "kind": kind, "key": kind, "severity": "error",
            "title": f"{_people(len(rows))} " + title,
            "detail": detail,
            "n_blocked": len(rows),
            "employees": [{"employee_key": r["employee_key"],
                           "file_number": r["file_number"], "name": r["name"]}
                          for r in rows],
            "remedies": [{
                "action": "set_supervisor",
                "label": "Point them at somebody else",
                "note": "Writes a Supervisor Map row, which overrides ADP and "
                        "survives the next import.",
                "fields": {"supervisor_id": "",
                           "file_numbers": [r["file_number"] for r in rows]},
            }],
        })
    return out


def _email_group(conn: sqlite3.Connection) -> list[dict]:
    """Missing work email - a warning, and only ever fixed person by person."""
    people = [dict(r) for r in conn.execute(
        """
        SELECT e.employee_key, e.file_number,
               e.legal_last_name || ', ' || e.legal_first_name AS name
          FROM ADP_Concur_Employees e
         WHERE e.row_state <> 'deleted' AND (e.work_email IS NULL OR e.work_email = '')
         ORDER BY e.legal_last_name
        """)]
    if not people:
        return []
    return [{
        "kind": "work_email", "key": "work_email", "severity": "warning",
        "title": f"{_people(len(people))} {_have(len(people))} no work email "
                 "address in ADP",
        "detail": "Fixed one at a time - open the record and type the address. "
                  "Doing it here rather than in ADP is a stopgap: the next "
                  "import will overwrite it with whatever ADP still says.",
        "n_blocked": len(people),
        "employees": [{"employee_key": p["employee_key"],
                       "file_number": p["file_number"], "name": p["name"]}
                      for p in people],
        "remedies": [],
    }]


def ADP_Concur_fix_plan(conn: sqlite3.Connection) -> dict:
    """
    Everything wrong, grouped by what would fix it, worst first.

    Errors before warnings, and within each, whatever blocks the most people
    first - so working down the list clears the extract as fast as it can be
    cleared.
    """
    groups = (_missing_people(conn) + _unmapped_groups(conn) + _chain_groups(conn)
              + _supervisor_map_groups(conn) + _login_group(conn)
              + _email_group(conn))
    groups.sort(key=lambda g: (0 if g["severity"] == "error" else 1, -g["n_blocked"]))
    return {
        "groups": groups,
        "totals": {
            "groups": len(groups),
            "errors": sum(1 for g in groups if g["severity"] == "error"),
            "blocked": sum(g["n_blocked"] for g in groups if g["severity"] == "error"),
        },
    }


# ----------------------------------------------------------------- applying


def _counts(conn: sqlite3.Connection) -> dict:
    row = conn.execute(
        "SELECT COALESCE(SUM(severity = 'error'), 0), "
        "COALESCE(SUM(severity = 'warning'), 0) FROM ADP_Concur_Exceptions"
    ).fetchone()
    return {"errors": row[0], "warnings": row[1]}


def ADP_Concur_apply_fix(conn: sqlite3.Connection, action: str,
                         fields: dict) -> dict:
    """
    Apply one remedy, then re-derive.

    Everything comes back with the error and warning counts before and after,
    because that is the only honest report on whether a fix worked - and it
    catches the case where a fix resolves two problems and creates a third,
    which creating a person from guesses can easily do.
    """
    before = _counts(conn)
    cfg = load_config()
    detail: dict = {}

    if action == "create_employee":
        detail = _apply_create_employee(conn, fields)
    elif action == "set_supervisor":
        detail = _apply_set_supervisor(conn, fields)
    elif action == "add_map_row":
        detail = _apply_add_map_row(conn, fields)
    elif action == "delete_map_row":
        detail = _apply_delete_map_row(conn, fields)
    elif action == "set_login_sources":
        detail = _apply_login_sources(cfg, fields)
        cfg = load_config()
    elif action == "set_employee_fields":
        detail = _apply_employee_fields(conn, fields)
    elif action in ("exclude_employee", "include_employee"):
        detail = _apply_include(conn, fields, action == "include_employee")
    else:
        raise ValueError(f"No such fix: {action}")

    conn.commit()
    derived = ADP_Concur_derive(conn, cfg)
    after = _counts(conn)

    # What is still wrong with whoever the fix touched, so a created person's
    # own new problems are reported rather than discovered later.
    remaining = []
    for fn in detail.get("touched", []):
        remaining += [dict(r) for r in conn.execute(
            "SELECT file_number, employee_name, severity, field, message "
            "FROM ADP_Concur_Exceptions WHERE file_number = ?", (fn,))]

    return {"action": action, "detail": detail, "derive": derived,
            "before": before, "after": after,
            "resolved": before["errors"] - after["errors"],
            "remaining": remaining}


def _apply_create_employee(conn: sqlite3.Connection, fields: dict) -> dict:
    from ADP_Concur_Db import ADP_COLUMNS

    allowed = {c for _, c in ADP_COLUMNS}
    values = {k: _s(v) for k, v in fields.items() if k in allowed}
    file_number = values.get("file_number", "")
    if not file_number:
        raise ValueError("A File Number is required - it is what the people "
                         "reporting to this person are pointing at.")
    if conn.execute("SELECT 1 FROM ADP_Concur_Employees WHERE file_number = ?",
                    (file_number,)).fetchone():
        raise ValueError(f"File Number {file_number} is already in the load.")

    conn.execute(
        f"INSERT INTO ADP_Concur_Employees ({', '.join(values)}, source, row_state) "
        f"VALUES ({', '.join('?' for _ in values)}, 'manual', 'new')",
        list(values.values()))
    reports = [r[0] for r in conn.execute(
        "SELECT file_number FROM ADP_Concur_Employees WHERE supervisor_id = ?",
        (file_number,))]
    return {"created": file_number, "reports_resolved": len(reports),
            "touched": [file_number]}


def _apply_set_supervisor(conn: sqlite3.Connection, fields: dict) -> dict:
    """
    Write a Supervisor Map row for each named employee.

    The map rather than the employee's own Supervisor ID on purpose: the map
    is an override the derive always honours, and the next ADP import replaces
    the employee's columns but leaves the map alone - so the correction sticks.
    """
    file_numbers = [_s(f) for f in (fields.get("file_numbers") or []) if _s(f)]
    supervisor_id = _s(fields.get("supervisor_id"))
    note = _s(fields.get("note"))
    if not file_numbers:
        raise ValueError("Nobody named.")
    if supervisor_id and not conn.execute(
        "SELECT 1 FROM ADP_Concur_Employees WHERE file_number = ? "
        "AND row_state <> 'deleted'", (supervisor_id,)).fetchone():
        raise ValueError(f"{supervisor_id} is not in this load either - create "
                         "them first, or leave the supervisor blank.")

    sup = conn.execute(
        "SELECT legal_last_name || ', ' || legal_first_name AS name "
        "FROM ADP_Concur_Employees WHERE file_number = ?",
        (supervisor_id,)).fetchone() if supervisor_id else None

    for fn in file_numbers:
        emp = conn.execute(
            "SELECT legal_last_name || ', ' || legal_first_name AS name "
            "FROM ADP_Concur_Employees WHERE file_number = ?", (fn,)).fetchone()
        conn.execute(
            "INSERT INTO ADP_Concur_SupervisorMap "
            "(file_number, employee_name, supervisor_name, supervisor_id, note, row_state) "
            "VALUES (?, ?, ?, ?, ?, 'new') "
            "ON CONFLICT (file_number) DO UPDATE SET "
            "supervisor_id = excluded.supervisor_id, "
            "supervisor_name = excluded.supervisor_name, "
            "note = excluded.note, row_state = 'modified'",
            (fn, emp["name"] if emp else "", sup["name"] if sup else "",
             supervisor_id, note or "Set in the app"))
    return {"file_numbers": file_numbers, "supervisor_id": supervisor_id,
            "touched": file_numbers}


def _apply_add_map_row(conn: sqlite3.Connection, fields: dict) -> dict:
    name = _s(fields.get("map"))
    if name not in MAP_TABLES:
        raise ValueError(f"No such map: {name}")
    table, columns = MAP_TABLES[name]
    values = {c: _s(fields.get(c)) for c in columns if c in fields}
    if not values:
        raise ValueError("Nothing to save.")
    conn.execute(
        f"INSERT OR REPLACE INTO {table} ({', '.join(values)}, row_state) "
        f"VALUES ({', '.join('?' for _ in values)}, 'new')", list(values.values()))
    return {"map": name, "row": values, "touched": []}


def _apply_delete_map_row(conn: sqlite3.Connection, fields: dict) -> dict:
    """
    Remove one mapping row, keyed on the map's own key column.

    Only the first column of each map is accepted as the key, which is the one
    the table is unique on - so this can delete the row somebody meant and not
    a set of rows that happen to share a value.
    """
    name = _s(fields.get("map"))
    if name not in MAP_TABLES:
        raise ValueError(f"No such map: {name}")
    table, columns = MAP_TABLES[name]
    key_column = columns[0]
    key = _s(fields.get(key_column))
    if not key:
        raise ValueError(f"A {key_column.replace('_', ' ')} is needed to say "
                         "which row to remove.")
    cur = conn.cursor()
    cur.execute(f"DELETE FROM {table} WHERE {key_column} = ?", (key,))
    if not cur.rowcount:
        raise ValueError(f"No {name} map row for {key}.")
    return {"map": name, "removed": cur.rowcount, key_column: key,
            "touched": [key] if name == "supervisor" else []}


def _apply_login_sources(cfg: dict, fields: dict) -> dict:
    sources = [s.strip() for s in _s(fields.get("sources")).split(",") if s.strip()]
    known = {"work_email", "personal_email", "file_number"}
    unknown = [s for s in sources if s not in known]
    if unknown:
        raise ValueError(f"Not an employee field: {', '.join(unknown)}. "
                         f"Use {', '.join(sorted(known))}.")
    if not sources:
        raise ValueError("At least one source is needed.")
    cfg.setdefault("login_id", {})["sources"] = sources
    if "bare_domain" in fields:
        cfg["login_id"]["bare_domain"] = _s(fields.get("bare_domain"))
    save_config(cfg)
    return {"sources": sources, "touched": []}


def _apply_employee_fields(conn: sqlite3.Connection, fields: dict) -> dict:
    from ADP_Concur_Db import ADP_COLUMNS

    allowed = {c for _, c in ADP_COLUMNS} - {"file_number"}
    file_number = _s(fields.get("file_number"))
    values = {k: _s(v) for k, v in fields.items() if k in allowed}
    if not file_number or not values:
        raise ValueError("A File Number and at least one field are needed.")
    conn.execute(
        f"UPDATE ADP_Concur_Employees SET {', '.join(f'{k} = ?' for k in values)}, "
        "row_state = CASE WHEN row_state = 'new' THEN 'new' ELSE 'modified' END, "
        "modified_at = datetime('now') WHERE file_number = ?",
        list(values.values()) + [file_number])
    return {"file_number": file_number, "fields": list(values),
            "touched": [file_number]}


def _apply_include(conn: sqlite3.Connection, fields: dict,
                   include: bool) -> dict:
    """
    Leave somebody out of the load, or put them back.

    A soft delete, like the one on the employee editor: the row stays and is
    simply not written, so a leaver excluded today can be restored tomorrow
    without re-importing anything.
    """
    file_number = _s(fields.get("file_number"))
    if not file_number:
        raise ValueError("Nobody named.")
    conn.execute(
        "UPDATE ADP_Concur_Employees SET row_state = ?, modified_at = datetime('now') "
        "WHERE file_number = ?",
        ("modified" if include else "deleted", file_number))
    return {"file_number": file_number, "included": include,
            "touched": [file_number] if include else []}


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="List what is wrong, grouped by what would fix it.")
    ap.add_argument("--db", default=None, help=f"SQLite path (default {DEFAULT_DB_PATH})")
    ap.add_argument("--errors-only", action="store_true")
    args = ap.parse_args(argv)

    conn = connect(args.db)
    print(f"Database: {resolve_db_path(args.db)}")
    plan = ADP_Concur_fix_plan(conn)
    t = plan["totals"]
    print(f"{t['groups']} thing(s) to fix, {t['errors']} of them blocking "
          f"{t['blocked']} employee(s)\n")
    for g in plan["groups"]:
        if args.errors_only and g["severity"] != "error":
            continue
        print(f"[{g['severity']}] {g['title']}  ({g['n_blocked']} affected)")
        print(f"    {g['detail']}")
        for e in g["employees"][:6]:
            print(f"      {e['file_number']}  {e['name']}")
        if len(g["employees"]) > 6:
            print(f"      ... and {len(g['employees']) - 6} more")
        for r in g["remedies"]:
            print(f"    -> {r['label']}  ({r['action']})")
        print()
    conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
