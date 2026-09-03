"""
ADP_Concur_Map - the derivation engine.

Two jobs.

First, rebuild the columns Kelly derives with lookups in the workbook (AC..AM
on sheet "1") from the six map tables, so a mapping change is a row edit rather
than a spreadsheet edit. Every formula in the workbook has a function here with
the formula it replaces quoted above it.

Second, build the 305, 350 and 360 records from those derived values. The field
maps below are exactly what the workbook's 305/350/360 tabs point at, read out
of the cell formulas - so this file *is* the specification of the extract, and
a column that is not in a field map is one the template leaves empty.

Everything that could reasonably differ between runs - the Login ID rule, the
default password, which record types terminated people appear in - lives in
ADP_Concur_Config.json rather than in code.
"""
from __future__ import annotations

import json
import re
import sqlite3
from datetime import date, datetime
from pathlib import Path

CONFIG_PATH = Path(__file__).parent / "ADP_Concur_Config.json"

# What the workbook writes when a lookup misses. Kept identical so a value that
# fails here is recognisable to anyone who has been working in the spreadsheet.
UNMAPPED = {
    "org_unit_1": "BU Description Not Mapped",
    "org_unit_2": "Home Department Code Not Mapped",
    "concur_profile": "Salary Code Not Mapped",
    "travel_profile": "Salary Code Not Mapped",
    "legal_country": "Legal Country Not Mapped",
    "locale_code": "ADP and BU Language Code Not Mapped",
    "reimbursement_currency": "BU Currency Not Mapped",
    "concur_status": "Position Status Not Mapped",
}

DEFAULT_CONFIG = {
    # --------------------------------------------------------- login id
    # Concur Login IDs must be unique across every Concur entity, not just
    # yours, which is why the raw work email will not load as-is. The rule is
    # built here: take the first source that has a value, then apply the
    # suffix. A suffix of '.fmg' turns 'a.user@onebarnes.com' into
    # 'a.user@onebarnes.com.fmg', which is the shape SAP normally asks for;
    # some entities use a dedicated domain instead - set 'replace_domain'.
    "login_id": {
        # Work email only by default, matching the workbook. Add
        # 'personal_email' or 'file_number' to fall back rather than leaving
        # the 30-odd people with no work address out of the load.
        "sources": ["work_email"],
        "suffix": "",
        "replace_domain": "",
        "prefix": "",
        # Off by default so the extract reproduces the workbook exactly. ADP
        # writes 23 of these addresses with capitals ('TAllen2@...'); turn this
        # on once you know whether Concur cares.
        "lowercase": False,
        # Used when the chosen source is the file number, which has no '@'.
        "bare_domain": "",
    },
    "password": "Welcome01",
    # Which people reach each record type. 'active' means concur_status = 'Y'.
    "scope": {"305": "all", "350": "active", "360": "active"},
    # Flat file shape.
    "extract": {
        "delimiter": ",",
        "quote": "minimal",          # 'minimal' | 'all' | 'none'
        "line_ending": "\r\n",
        "encoding": "utf-8",
        "order": "by_type",          # 'by_type' (305s, 350s, 360s) | 'by_employee'
        "file_name": "FMG_Concur_Employee_{stamp}.txt",
        # A selection-scoped file gets its own name so a pilot load is never
        # mistaken for the full company sitting in the same pickup folder.
        "selection_file_name": "FMG_Concur_Employee_Selection_{stamp}.txt",
        "outbound_dir": "",          # blank = output/adp_concur/ beside this file
    },
    # Rows the extract refuses to write, and why. Turn one off to let it
    # through and see what Concur says.
    "block_on": {
        "unmapped": True,            # any '... Not Mapped' value
        "missing_login_id": True,
        "missing_employee_id": True,
        "broken_supervisor": True,   # an approver Concur cannot resolve
    },
}


def load_config(path: Path | str | None = None) -> dict:
    """Config with the defaults filled in for anything the file leaves out."""
    p = Path(path) if path else CONFIG_PATH
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))  # deep copy
    if p.exists():
        try:
            stored = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return cfg
        for key, value in stored.items():
            if isinstance(value, dict) and isinstance(cfg.get(key), dict):
                cfg[key].update(value)
            else:
                cfg[key] = value
    return cfg


def save_config(cfg: dict, path: Path | str | None = None) -> Path:
    p = Path(path) if path else CONFIG_PATH
    p.write_text(json.dumps(cfg, indent=2) + "\n", encoding="utf-8")
    return p


# ------------------------------------------------------------------ helpers


def _s(value) -> str:
    """Everything out of a cell or a column as a trimmed string."""
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _yyyymmdd(value) -> str:
    """
    TEXT(<date>,"yyyymmdd").

    Dates arrive as datetimes from openpyxl and as ISO strings out of SQLite,
    so both are handled. Anything unparseable comes back as given rather than
    being dropped, because a term date that is not a date is worth seeing.
    """
    if value is None or value == "":
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y%m%d")
    text = str(value).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y%m%d")
        except ValueError:
            continue
    return text


def column_ref_to_index(ref: str) -> int:
    """'A' -> 1, 'AB' -> 28. The field maps read better in template letters."""
    n = 0
    for ch in ref.strip().upper():
        n = n * 26 + (ord(ch) - 64)
    return n


# --------------------------------------------------------------- the lookups
#
# One function per derived column, each quoting the workbook formula it stands
# in for. `maps` is the dict returned by load_maps() - the six lookup tables
# indexed for lookup rather than re-queried per employee.


def load_maps(conn: sqlite3.Connection) -> dict:
    """Read the six map tables into dictionaries, once per derive run."""
    maps: dict = {}

    maps["status"] = {
        _s(r["position_status"]): _s(r["concur_status"])
        for r in conn.execute("SELECT * FROM ADP_Concur_StatusMap")
    }
    maps["country"] = {
        _s(r["adp_country"]): _s(r["concur_country"])
        for r in conn.execute("SELECT * FROM ADP_Concur_CountryMap")
    }
    maps["language"] = {
        _s(r["language_desc"]): _s(r["adp_language"])
        for r in conn.execute("SELECT * FROM ADP_Concur_LanguageMap")
    }
    maps["salary"] = {
        _s(r["pay_grade_code"]): (_s(r["expense_map"]), _s(r["travel_map"]))
        for r in conn.execute("SELECT * FROM ADP_Concur_SalaryMap")
    }
    maps["supervisor"] = {
        _s(r["file_number"]): _s(r["supervisor_id"])
        for r in conn.execute("SELECT * FROM ADP_Concur_SupervisorMap")
    }

    # The Org Map is read three ways: by business unit (for org unit 1, the
    # default language and the currency) and by home department code (for org
    # unit 2). VLOOKUP takes the first match, so the first row of each key wins.
    by_bu: dict = {}
    by_dept: dict = {}
    for r in conn.execute("SELECT * FROM ADP_Concur_OrgMap ORDER BY map_key"):
        bu = _s(r["business_unit_desc"])
        dept = _s(r["home_department_code"])
        if bu and bu not in by_bu:
            by_bu[bu] = {"org_unit_1": _s(r["org_unit_1"]),
                         "default_language": _s(r["default_language"]),
                         "currency": _s(r["currency"])}
        if dept and dept not in by_dept:
            by_dept[dept] = _s(r["org_unit_2"])
    maps["org_by_bu"] = by_bu
    maps["org_by_dept"] = by_dept
    return maps


def ADP_Concur_supervisor_id(emp: dict, maps: dict) -> str:
    """
    AC: =IF(ISBLANK(H5),VLOOKUP(P5,'Supervisor Map'!A:D,4,FALSE),MID(H5,4,6))

    ADP writes the supervisor as a three-character company prefix plus the
    file number - 'KSY211756' - so the middle six characters are the ID Concur
    wants. A mapped blank is deliberate and means top of the food chain.

    One deliberate difference from the workbook: **the Supervisor Map wins**,
    rather than only being consulted when ADP is blank. A table of exceptions
    that can fill a hole but cannot correct a wrong value is no use for the
    thing it is most needed for - somebody reporting to an ID that is not in
    the load - and correcting those is the whole point of the Fixes tab.

    On the workbook as it stands this changes nothing: of the seven map rows,
    only Moavero also has an ADP value, and both give 006769.
    """
    mapped = maps["supervisor"].get(_s(emp.get("file_number")))
    if mapped is not None:
        return mapped
    raw = _s(emp.get("supervisor_id_raw"))
    return raw[3:9] if raw else ""


def ADP_Concur_org_unit_1(emp: dict, maps: dict) -> str:
    """AD: =IFERROR(VLOOKUP(T5,'Org Map'!A:E,4),"BU Description Not Mapped")"""
    hit = maps["org_by_bu"].get(_s(emp.get("business_unit_desc")))
    return hit["org_unit_1"] if hit else UNMAPPED["org_unit_1"]


def ADP_Concur_org_unit_2(emp: dict, maps: dict) -> str:
    """
    AE: =IFERROR(VLOOKUP(V5,'Org Map'!C:E,3),"Home Department Code Not Mapped")

    The workbook leaves the fourth VLOOKUP argument off, so Excel does an
    approximate match and can silently return the row above the one it wanted.
    This is an exact match on the home department code - a code the Org Map
    does not carry is reported rather than mapped to its neighbour.
    """
    hit = maps["org_by_dept"].get(_s(emp.get("home_department_code")))
    return hit if hit else UNMAPPED["org_unit_2"]


def ADP_Concur_concur_profile(emp: dict, maps: dict) -> str:
    """AF: =IFERROR(VLOOKUP(AA5,'Salary Map'!A:D,3,FALSE),"Salary Code Not Mapped")"""
    hit = maps["salary"].get(_s(emp.get("pay_grade_code")))
    return hit[0] if hit else UNMAPPED["concur_profile"]


def ADP_Concur_travel_profile(emp: dict, maps: dict) -> str:
    """AG: =IFERROR(VLOOKUP(AA5,'Salary Map'!A:D,4,FALSE),"Salary Code Not Mapped")"""
    hit = maps["salary"].get(_s(emp.get("pay_grade_code")))
    return hit[1] if hit else UNMAPPED["travel_profile"]


def ADP_Concur_legal_country(emp: dict, maps: dict) -> str:
    """AH: =IFERROR(VLOOKUP(Z5,'Country Map'!A:B,2,FALSE),"Legal Country Not Mapped")"""
    hit = maps["country"].get(_s(emp.get("legal_country_code")))
    return hit if hit else UNMAPPED["legal_country"]


def ADP_Concur_locale_code(emp: dict, maps: dict, legal_country: str) -> str:
    """
    AI: =IFERROR(VLOOKUP(X5,'Language Map'!A:C,3,FALSE),
           IFERROR(VLOOKUP(T5,'Org Map'!A:F,6,FALSE),
             "ADP and BU Language Code Not Mapped"))&AH5

    The language stem and the country are concatenated: 'en_' & 'US'. ADP's own
    language description wins; where ADP has none, the business unit's default
    language is used. Note the workbook concatenates the country even onto the
    not-mapped text, and so does this - it keeps the two readable together.
    """
    stem = maps["language"].get(_s(emp.get("language_desc")))
    if not stem:
        hit = maps["org_by_bu"].get(_s(emp.get("business_unit_desc")))
        stem = hit["default_language"] if hit else ""
    if not stem:
        stem = UNMAPPED["locale_code"]
    return f"{stem}{legal_country}"


def ADP_Concur_reimbursement_currency(emp: dict, maps: dict) -> str:
    """AJ: =IFERROR(VLOOKUP(T5,'Org Map'!A:G,7),"BU Currency Not Mapped")"""
    hit = maps["org_by_bu"].get(_s(emp.get("business_unit_desc")))
    return hit["currency"] if hit else UNMAPPED["reimbursement_currency"]


def ADP_Concur_preferred_name(emp: dict) -> str:
    """AK: =IF(ISBLANK(E5),"",TRIM(E5&" "&F5))"""
    first = _s(emp.get("preferred_first_name"))
    if not first:
        return ""
    return f"{first} {_s(emp.get('preferred_last_name'))}".strip()


def ADP_Concur_status(emp: dict, maps: dict) -> str:
    """AL: =IFERROR(VLOOKUP(O5,'Status Map'!A:B,2,FALSE),"Position Status Not Mapped")"""
    hit = maps["status"].get(_s(emp.get("position_status")))
    return hit if hit else UNMAPPED["concur_status"]


def ADP_Concur_term_date(emp: dict, concur_status: str) -> str:
    """AM: =IF(AL5="N",TEXT(M5,"yyyymmdd"),"")"""
    if concur_status != "N":
        return ""
    return _yyyymmdd(emp.get("termination_date"))


def ADP_Concur_login_id(emp: dict, cfg: dict) -> str:
    """
    Not in the workbook - this is the piece Kelly's email is asking about.

    Concur Login IDs are unique across every entity on the platform, so the
    plain work email address usually collides with one that already exists
    somewhere. The rule is configurable rather than guessed:

      sources         first field with a value wins
      replace_domain  swap the part after '@' for this
      suffix          appended whole, e.g. '.fmg' -> 'a@onebarnes.com.fmg'
      prefix          prepended whole
      bare_domain     '@x' added when the source has no '@' (a file number)

    Out of the box nothing is applied, so the login IDs come out as the work
    email exactly like the workbook - set the rule once you know which shape
    SAP wants and the whole extract follows it.
    """
    rule = cfg.get("login_id", {})
    value = ""
    for src in rule.get("sources") or ["work_email"]:
        value = _s(emp.get(src))
        if value:
            break
    if not value:
        return ""

    if "@" not in value and rule.get("bare_domain"):
        domain = rule["bare_domain"].lstrip("@")
        value = f"{value}@{domain}"
    elif "@" in value and rule.get("replace_domain"):
        domain = rule["replace_domain"].lstrip("@")
        value = f"{value.split('@', 1)[0]}@{domain}"

    value = f"{rule.get('prefix', '')}{value}{rule.get('suffix', '')}"
    return value.lower() if rule.get("lowercase", False) else value


# ------------------------------------------------------------------- derive


DERIVED_FIELDS = ["supervisor_id", "org_unit_1", "org_unit_2", "concur_profile",
                  "travel_profile", "legal_country", "locale_code",
                  "reimbursement_currency", "preferred_name", "concur_status",
                  "term_date", "login_id"]


def derive_one(emp: dict, maps: dict, cfg: dict) -> dict:
    """Every derived value for one employee, in the workbook's own order."""
    legal_country = ADP_Concur_legal_country(emp, maps)
    concur_status = ADP_Concur_status(emp, maps)
    return {
        "supervisor_id": ADP_Concur_supervisor_id(emp, maps),
        "org_unit_1": ADP_Concur_org_unit_1(emp, maps),
        "org_unit_2": ADP_Concur_org_unit_2(emp, maps),
        "concur_profile": ADP_Concur_concur_profile(emp, maps),
        "travel_profile": ADP_Concur_travel_profile(emp, maps),
        "legal_country": legal_country,
        "locale_code": ADP_Concur_locale_code(emp, maps, legal_country),
        "reimbursement_currency": ADP_Concur_reimbursement_currency(emp, maps),
        "preferred_name": ADP_Concur_preferred_name(emp),
        "concur_status": concur_status,
        "term_date": ADP_Concur_term_date(emp, concur_status),
        "login_id": ADP_Concur_login_id(emp, cfg),
    }


def ADP_Concur_derive(conn: sqlite3.Connection, cfg: dict | None = None,
                      commit: bool = True) -> dict:
    """
    Recompute every derived column and rebuild the exception list.

    Cheap enough to run after any change - a map edit, a new employee, a new
    Login ID rule - and that is how the app uses it, so the Employees tab and
    the extract can never disagree about what a lookup returns.
    """
    cfg = cfg or load_config()
    maps = load_maps(conn)
    rows = [dict(r) for r in conn.execute(
        "SELECT * FROM ADP_Concur_Employees WHERE row_state <> 'deleted'")]

    sets = ", ".join(f"{f} = ?" for f in DERIVED_FIELDS)
    updates = []
    exceptions = []

    for emp in rows:
        derived = derive_one(emp, maps, cfg)
        updates.append([derived[f] for f in DERIVED_FIELDS] + [emp["employee_key"]])
        exceptions.extend(check_one(emp, derived, cfg))

    cur = conn.cursor()
    cur.executemany(
        f"UPDATE ADP_Concur_Employees SET {sets}, derived_at = datetime('now') "
        "WHERE employee_key = ?", updates)

    # The chain checks have to run after the supervisor_id updates are in the
    # connection - they walk the table, not the dicts. Uncommitted is fine:
    # it is the same connection.
    exceptions.extend(check_hierarchy(conn, cfg))

    cur.execute("DELETE FROM ADP_Concur_Exceptions")
    cur.executemany(
        "INSERT INTO ADP_Concur_Exceptions "
        "(employee_key, file_number, employee_name, severity, field, message) "
        "VALUES (?, ?, ?, ?, ?, ?)", exceptions)

    if commit:
        conn.commit()

    errors = sum(1 for e in exceptions if e[3] == "error")
    return {"employees": len(rows), "exceptions": len(exceptions),
            "errors": errors, "warnings": len(exceptions) - errors}


def check_hierarchy(conn: sqlite3.Connection, cfg: dict) -> list[tuple]:
    """
    What is wrong with the supervisor chain, as exception rows.

    A supervisor_id that names nobody in the load is the one that actually
    stops a load: it is written into the 305 as the Expense Report Approver
    and into the 360 twice more, and Concur will not accept an approver it
    cannot resolve. It is an error by default - `block_on.broken_supervisor`
    turns it into a warning if you would rather send it and see.

    A supervisor who *is* in the load but is being held out of it is the same
    problem arriving a step later, so it is called out separately rather than
    left to be discovered in Concur.
    """
    from ADP_Concur_Hierarchy import ADP_Concur_hierarchy_problems

    block = cfg.get("block_on", {})
    severity = "error" if block.get("broken_supervisor", True) else "warning"
    problems = ADP_Concur_hierarchy_problems(conn)
    out = []

    def add(row, sev, message):
        out.append((row["employee_key"], row["file_number"], row["name"],
                    sev, "supervisor_chain", message))

    for r in problems["broken"]:
        named = _s(r.get("reports_to_legal_name"))
        add(r, severity,
            f"Supervisor {r['supervisor_id']} is not in this load"
            + (f" - ADP says \"{named}\"" if named else "")
            + ". Concur cannot resolve the approver.")

    for r in problems["self_led"]:
        add(r, severity, "Reports to themselves - Concur will not accept a "
                         "record as its own approver.")

    for r in problems["cycles"]:
        add(r, severity,
            f"Inside a supervisor loop - following {r['supervisor_id']} upwards "
            "comes back here rather than reaching the top.")

    # Supervisors who exist but will not be in the file.
    for r in conn.execute(
        """
        SELECT e.employee_key, e.file_number,
               e.legal_last_name || ', ' || e.legal_first_name AS name,
               e.supervisor_id,
               s.legal_last_name || ', ' || s.legal_first_name AS sup_name,
               s.row_state AS sup_state, s.include_305 AS sup_305
          FROM ADP_Concur_Employees e
          JOIN ADP_Concur_Employees s ON s.file_number = e.supervisor_id
         WHERE e.row_state <> 'deleted' AND e.supervisor_id <> ''
           AND (s.row_state = 'deleted' OR s.include_305 = 0)
        """
    ):
        out.append((r["employee_key"], r["file_number"], r["name"], "warning",
                    "supervisor_chain",
                    f"Supervisor {r['sup_name']} ({r['supervisor_id']}) is being "
                    "held out of the 305, so the approver will not exist in Concur "
                    "when this record loads."))

    return out


def check_one(emp: dict, derived: dict, cfg: dict) -> list[tuple]:
    """
    Everything wrong with one employee, as ADP_Concur_Exceptions rows.

    'error' is something the extract will refuse to write; 'warning' is
    something worth a look that still loads. Kelly's email names three of these
    directly - bad or incomplete ADP data, terminated people whose mapping no
    longer resolves, and the missing supervisors.
    """
    key = emp["employee_key"]
    fn = _s(emp.get("file_number"))
    name = f"{_s(emp.get('legal_last_name'))}, {_s(emp.get('legal_first_name'))}".strip(", ")
    out = []
    block = cfg.get("block_on", {})
    terminated = derived["concur_status"] == "N"

    def add(sev, field, message):
        out.append((key, fn, name, sev, field, message))

    if not fn:
        add("error" if block.get("missing_employee_id", True) else "warning",
            "file_number", "No File Number - Concur has no Employee ID to key on.")

    if not derived["login_id"]:
        add("error" if block.get("missing_login_id", True) else "warning",
            "login_id", "No Login ID - ADP has neither a work nor a personal "
                        "email address, and the rule has no fallback.")

    if not _s(emp.get("work_email")):
        add("warning", "work_email", "No work email address in ADP.")

    for field, text in UNMAPPED.items():
        if derived.get(field) and text in str(derived[field]):
            sev = "warning" if terminated else (
                "error" if block.get("unmapped", True) else "warning")
            add(sev, field, f"{text}"
                            + (" (terminated - ignorable)" if terminated else ""))

    if not derived["supervisor_id"] and not terminated:
        add("warning", "supervisor_id",
            "No supervisor. ADP has no Supervisor ID and the Supervisor Map "
            "has no entry - add one, or leave it if this is the top of the "
            "food chain.")

    if terminated and not derived["term_date"]:
        add("warning", "term_date",
            "Inactive in Concur but ADP has no Termination Date.")

    if (emp.get("duplicate_rows") or 1) > 1:
        add("warning", "duplicate_rows",
            _s(emp.get("duplicate_note")) or
            "ADP sent more than one row for this File Number.")

    if _s(emp.get("legal_country_code")) not in ("USA", "US", ""):
        add("warning", "legal_country_code",
            f"Non-US employee ({_s(emp.get('legal_country_code'))}).")

    return out


# ----------------------------------------------------------- record building
#
# Read straight out of the workbook's 305/350/360 tabs. ('const', x) is a value
# the template hard-codes; ('field', x) is a column of ADP_Concur_Employees,
# raw or derived. A position that is not here is written empty, which is what
# the template does.

FIELD_MAP: dict[str, dict[str, tuple[str, str]]] = {
    "305": {
        "A":  ("const", "305"),
        "B":  ("field", "legal_first_name"),
        "C":  ("field", "middle_initial"),
        "D":  ("field", "legal_last_name"),
        "E":  ("field", "file_number"),              # Employee ID
        "F":  ("field", "login_id"),                 # Login ID
        "G":  ("config", "password"),
        "H":  ("field", "work_email"),               # Email Address
        "I":  ("field", "locale_code"),
        "L":  ("field", "org_unit_1"),               # Ledger Code
        "M":  ("field", "reimbursement_currency"),
        "O":  ("field", "concur_status"),            # Active (Y/N)
        "P":  ("field", "org_unit_1"),
        "Q":  ("field", "org_unit_2"),
        "W":  ("field", "pay_grade_code"),           # Custom 2 Salary Code
        "X":  ("field", "travel_profile"),           # Custom 3 Expense Profile
        "Z":  ("field", "org_unit_2"),               # Custom 5 Department
        "AB": ("field", "term_date"),                # Custom 7 Term Date
        "AC": ("field", "preferred_name"),           # Custom 8 Preferred Name
        "AP": ("field", "pay_grade_code"),           # Custom 21 Expense Group
        "BG": ("field", "supervisor_id"),            # Expense Report Approver
        "BK": ("const", "Y"),                        # Expense User
        "BL": ("const", "Y"),                        # Expense / Cash Advance Approver
        "BU": ("const", "Y"),                        # Invoice User
        "BV": ("const", "Y"),                        # Invoice Approver
        "CE": ("const", "Y"),                        # Future Use 2
        "CH": ("const", "Y"),                        # Travel Wizard User
    },
    "350": {
        "A": ("const", "350"),
        "B": ("field", "file_number"),
        "R": ("field", "travel_profile"),            # Travel Class Name
        "T": ("field", "org_unit_1"),                # Org Unit / Division
    },
    "360": {
        "A":  ("const", "360"),
        "B":  ("field", "file_number"),
        "C":  ("const", "Y"),                        # Invoice User Role
        "D":  ("const", "Y"),                        # Invoice Approver Role
        "I":  ("const", "Y"),                        # Purchase Request User
        "J":  ("const", "Y"),                        # Purchase Request Approver
        "R":  ("field", "supervisor_id"),            # Default PR Approver
        "S":  ("field", "supervisor_id"),            # Payment Approver
        "AC": ("const", "Y"),                        # Display Image In-line
        "AD": ("const", "Y"),                        # Auto Open Image
    },
}

# Used when the workbook layouts have not been captured, so the record still
# comes out the right width. Taken from the template Kelly is working from.
DEFAULT_WIDTHS = {"305": 137, "350": 67, "360": 35}


def layout_width(conn: sqlite3.Connection, record_type: str) -> int:
    """How many fields a record of this type carries."""
    row = conn.execute(
        "SELECT MAX(position) FROM ADP_Concur_Layouts WHERE record_type = ?",
        (record_type,)).fetchone()
    return (row and row[0]) or DEFAULT_WIDTHS.get(record_type, 0)


def build_record(emp: dict, record_type: str, width: int, cfg: dict) -> list[str]:
    """One record as a list of field values, padded to the layout width."""
    out = [""] * width
    for ref, (kind, value) in FIELD_MAP[record_type].items():
        idx = column_ref_to_index(ref) - 1
        if idx >= width:
            continue
        if kind == "const":
            out[idx] = value
        elif kind == "config":
            out[idx] = _s(cfg.get(value, ""))
        else:
            out[idx] = _s(emp.get(value))
    return out


def selected_employees(conn: sqlite3.Connection, record_type: str,
                       cfg: dict, keys: list[int] | None = None) -> list[dict]:
    """
    The people who belong in one record type.

    Four things decide it: the per-record include flag on the employee (so a
    single person can be held back without touching anything else), the
    configured scope - 'active' drops everyone the Status Map turns into 'N' -
    the exception list, which keeps blocking errors out of the file, and
    `keys`, which narrows the whole thing to a chosen set of people.

    `keys` is how a pilot load is built: pick one manager's organisation on the
    Hierarchy tab and the extract carries those people and nobody else. An
    empty list is not the same as None - it means nothing was chosen, and it
    produces an empty file rather than the whole company.
    """
    scope = (cfg.get("scope") or {}).get(record_type, "all")
    sql = (f"SELECT e.* FROM ADP_Concur_Employees e "
           f"WHERE e.row_state <> 'deleted' AND e.include_{record_type} = 1")
    args: list = []
    if scope == "active":
        sql += " AND e.concur_status = 'Y'"
    if keys is not None:
        if not keys:
            return []
        sql += f" AND e.employee_key IN ({','.join('?' * len(keys))})"
        args += list(keys)
    sql += (" AND NOT EXISTS (SELECT 1 FROM ADP_Concur_Exceptions x "
            "WHERE x.employee_key = e.employee_key AND x.severity = 'error')")
    sql += " ORDER BY e.legal_last_name, e.legal_first_name, e.file_number"
    return [dict(r) for r in conn.execute(sql, args)]


def build_records(conn: sqlite3.Connection, cfg: dict | None = None,
                  keys: list[int] | None = None) -> dict:
    """Every 305, 350 and 360 record the current data produces."""
    cfg = cfg or load_config()
    out = {}
    for record_type in ("305", "350", "360"):
        width = layout_width(conn, record_type)
        people = selected_employees(conn, record_type, cfg, keys)
        out[record_type] = [build_record(e, record_type, width, cfg) for e in people]
    return out


if __name__ == "__main__":
    import argparse

    from ADP_Concur_Db import DEFAULT_DB_PATH, connect

    ap = argparse.ArgumentParser(description="Rebuild the derived Concur values.")
    ap.add_argument("--db", default=None, help=f"SQLite path (default {DEFAULT_DB_PATH})")
    ap.add_argument("--show-config", action="store_true")
    args = ap.parse_args()

    config = load_config()
    if args.show_config:
        print(json.dumps(config, indent=2))
        raise SystemExit(0)

    c = connect(args.db)
    result = ADP_Concur_derive(c, config)
    print(f"{result['employees']} employees derived, "
          f"{result['errors']} error(s), {result['warnings']} warning(s)")
    for r in c.execute(
        "SELECT severity, field, COUNT(*) n FROM ADP_Concur_Exceptions "
        "GROUP BY severity, field ORDER BY severity, n DESC"
    ):
        print(f"  {r['severity']:<8} {r['field'] or '':<24} {r['n']}")
    c.close()
