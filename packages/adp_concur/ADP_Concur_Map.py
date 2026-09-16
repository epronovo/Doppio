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
    # The non-US tab's own wording, for the one lookup only it performs.
    "country_currency": "Country Currency Not Mapped",
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
        # Both 305 tabs now build the Login ID as =<email>&".new.uat", so this
        # belongs to the whole load rather than to one roster. It still reads
        # like a UAT-only value - worth confirming before a production run.
        "suffix": ".new.uat",
        "replace_domain": "",
        "prefix": "",
        # Off by default so the extract reproduces the workbook exactly. ADP
        # writes 23 of these addresses with capitals ('TAllen2@...'); turn this
        # on once you know whether Concur cares.
        "lowercase": False,
        # Used when the chosen source is the file number, which has no '@'.
        "bare_domain": "",
    },
    # Per-roster overrides, merged over login_id above. Empty because both
    # rosters currently agree; kept because they have disagreed before and the
    # mechanism is what let the extract match each tab while they did.
    "login_id_by_roster": {},
    "password": "Welcome01",
    # ------------------------------------------------- the 100 record
    # Import Settings, one per file, written as the first line. Seven fields,
    # all required, in this order - SAP's "Reviewing the Import Definition File
    # (Feed ID StandardEmployeeImport) - Import Settings (Record Type 100)".
    #
    #   error_threshold           SAP says enter 0
    #   password_generation       EMPID | LOGINID | TEXT | SSO
    #                             TEXT uses the password on the 305 record,
    #                             which is what both of these rosters carry
    #   existing_record_handling  REPLACE | UPDATE | WARN | IGNORE
    #                             UPDATE only writes the non-blank fields and
    #                             never overwrites an existing password, which
    #                             is the safe default for a repeated load;
    #                             REPLACE overwrites the record wholesale
    #   language_code             the language of any localised text in the file
    #   validate_expense_group    Y | N, SAP default Y
    #   validate_payment_group    Y | N, SAP default Y
    # These are not SAP's defaults - they are what FMG's own 100 record said on
    # the file that actually loaded in July ('305 360 import FMG 07.24.26.txt'):
    #
    #     100,0,SSO,UPDATE,EN,N,N
    #
    # SSO rather than TEXT, so the passwords on the 305 are ignored; EN rather
    # than en; and both group validations off. A proven load beats a
    # documentation default, so that is what ships.
    "import_settings": {
        "error_threshold": "0",
        "password_generation": "SSO",
        "existing_record_handling": "UPDATE",
        "language_code": "EN",
        "validate_expense_group": "N",
        "validate_payment_group": "N",
    },
    # Which record types each roster writes. In config rather than in code
    # because it is a real question: the July file carried 305 and 360 only,
    # and the workbook has since grown a 350 tab. Drop "350" here to go back to
    # what loaded before.
    "records_by_roster": {
        "us": ["305", "350", "360"],
        "non_us": ["305", "360"],
    },
    # Master switch per record type - turn one off and it is written nowhere,
    # in any file, whatever records_by_roster says. This exists because
    # records_by_roster only ever applied to a roster-scoped file: the plain
    # 'Preview' / 'Write extract' buttons write one combined file with no
    # roster and always carried all three types regardless of that setting.
    # This is the one switch both paths obey.
    "records_enabled": {"305": True, "350": True, "360": True},
    # Which people reach each record type. 'active' means concur_status = 'Y'.
    "scope": {"305": "all", "350": "active", "360": "active", "320": "all"},
    # Columns to write empty whatever the field map says, per record type.
    #
    # This exists because a column can be right by the workbook and still wrong
    # by the tenant. Concur's Run-18 result warned 120 times that Custom 5 -
    # then the Org Unit 2 department code, now Org Unit 1 - "could not be
    # resolved to an existing custom list item", and the July file that loaded
    # cleanly left Custom 5 blank on all 216 of its 305 records. Whether the
    # list is missing from the tenant or the column is simply not wanted is a
    # question for SAP, and this is how you answer it without editing code:
    # add "Z" to blank it, take it out to send it again.
    #
    # Nothing is blanked by default - the extract still reproduces the
    # workbook, and every departure from it stays a decision somebody made.
    "blank_fields": {"305": [], "350": [], "360": [], "320": []},
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
        # The 320 is a separate file by SAP's own rule - see
        # ADP_Concur_Export.ADP_Concur_export_320 - so it gets its own name
        # rather than sharing file_name above.
        "file_name_320": "FMG_Concur_UpdateID_{stamp}.txt",
    },
    # Rows the extract refuses to write, and why. Turn one off to let it
    # through and see what Concur says.
    "block_on": {
        "unmapped": True,            # any '... Not Mapped' value
        "missing_login_id": True,
        "missing_employee_id": True,
        "broken_supervisor": True,   # an approver Concur cannot resolve
        # Off, because the rule is inferred from Concur's rejections rather
        # than from a list we hold - see the check in employee_exceptions. Turn
        # it on to keep the 54 non-US people whose Org Unit 2 is a department
        # name out of the file instead of watching Concur refuse them.
        "org_unit_2_code": False,
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

    # Only the non-US roster reads these two: country -> locale, and
    # country -> currency code.
    maps["locale"] = {
        _s(r["country_code"]).upper(): _s(r["locale_code"])
        for r in conn.execute("SELECT * FROM ADP_Concur_LocaleMap")
    }
    maps["country_ref"] = {
        _s(r["country_code"]).upper(): _s(r["currency_code"])
        for r in conn.execute("SELECT * FROM ADP_Concur_CountryRef")
    }
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
    rule = dict(cfg.get("login_id", {}))
    rule.update((cfg.get("login_id_by_roster") or {}).get(
        _s(emp.get("roster")) or "us", {}))
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


def derive_non_us(emp: dict, maps: dict, cfg: dict) -> dict:
    """
    The same twelve values, for somebody on the non-US roster.

    A separate function rather than branches inside the US one, because it is
    genuinely a different derivation: those people are not in ADP at all. Their
    tab is maintained by hand in Concur's own terms, so most of what is a
    lookup for a US employee arrives already answered - the country is already
    a two-character code, the org unit is the ledger code, the status is
    already Y. What is left is three lookups, and each quotes the formula on
    the '305 Non US Non SE' tab that it stands in for.
    """
    country = _s(emp.get("legal_country_code")).upper()

    # I: =IFERROR(VLOOKUP(J5,'Language Map'!M:O,3,FALSE),"en_"&J5)
    locale = maps["locale"].get(country) or (f"en_{country}" if country else "")

    # M: =IFERROR(VLOOKUP(J5,'Country Map'!E:H,3,FALSE),"Country Currency Not Mapped")
    currency = maps["country_ref"].get(country) or UNMAPPED["country_currency"]

    # X: =IFERROR(VLOOKUP(W5,'Salary Map'!A:D,3,FALSE),"Salary Code Not Mapped")
    # Note this is column 3, the Expense Map - where the US 305 tab points its
    # equivalent at column 4, the Travel Map. Both are reproduced as written.
    grade = maps["salary"].get(_s(emp.get("pay_grade_code")))

    # P and AP: =L5. The ledger code is the org unit for this roster.
    ledger = _s(emp.get("ledger_code"))

    status = _s(emp.get("concur_status")).upper() or "Y"
    return {
        # Already a bare Concur Employee ID on this tab, so it passes straight
        # through - no three-character payroll prefix to strip.
        "supervisor_id": (maps["supervisor"].get(_s(emp.get("file_number")))
                          if _s(emp.get("file_number")) in maps["supervisor"]
                          else _s(emp.get("supervisor_id_raw"))),
        "org_unit_1": ledger,
        # Z: =Q5. Blank stays blank; the workbook's formula turns an empty cell
        # into a literal 0, which is an Excel artifact rather than a value.
        "org_unit_2": _s(emp.get("org_unit_2_raw")),
        "concur_profile": grade[0] if grade else UNMAPPED["concur_profile"],
        "travel_profile": grade[1] if grade else UNMAPPED["travel_profile"],
        "legal_country": country,
        "locale_code": locale,
        "reimbursement_currency": currency,
        "preferred_name": ADP_Concur_preferred_name(emp),
        "concur_status": status if status in ("Y", "N") else "Y",
        "term_date": "",
        "login_id": ADP_Concur_login_id(emp, cfg),
    }


def derive_one(emp: dict, maps: dict, cfg: dict) -> dict:
    """Every derived value for one employee, in the workbook's own order."""
    if _s(emp.get("roster")) == "non_us":
        return derive_non_us(emp, maps, cfg)
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

    # Approvers who are in the load but in the *other* file. Splitting the
    # extract by roster splits these apart: 12 people here report across the
    # line. Concur has to already know the approver when the record lands, so
    # the file holding the approver goes first - which is a warning about load
    # order rather than something wrong with the data.
    for r in conn.execute(
        """
        SELECT e.employee_key, e.file_number,
               e.legal_last_name || ', ' || e.legal_first_name AS name,
               e.roster, e.supervisor_id, s.roster AS sup_roster,
               s.legal_last_name || ', ' || s.legal_first_name AS sup_name
          FROM ADP_Concur_Employees e
          JOIN ADP_Concur_Employees s ON s.file_number = e.supervisor_id
         WHERE e.row_state <> 'deleted' AND s.row_state <> 'deleted'
           AND e.supervisor_id <> '' AND e.roster <> s.roster
        """
    ):
        out.append((r["employee_key"], r["file_number"], r["name"], "warning",
                    "supervisor_chain",
                    f"Approver {r['sup_name']} ({r['supervisor_id']}) is on the "
                    f"{r['sup_roster']} roster and this record is on the "
                    f"{r['roster']} one, so they land in different files. Load "
                    "the approver's file first."))

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

    # An unmapped value is an error whoever it belongs to.
    #
    # This used to downgrade terminated people to a warning, on the reasoning
    # that a leaver's stale BU is nobody's problem. Concur disagreed. Carla
    # Cunningham (007633) is terminated, carried 'BU Description Not Mapped'
    # through the downgrade into the file, and came back as a hard error:
    # "An employee fails to resolve to a valid EmployeeHierarchyService
    # configuration hierarchy node ... {Segment1=BU Description Not Mapped}".
    # Concur validates the hierarchy node before it looks at whether the person
    # is active, so being a leaver buys nothing. The status is still worth
    # saying out loud, because it tells you the cheap fix is usually to drop
    # the person from the load rather than to map a BU that no longer exists.
    for field, text in UNMAPPED.items():
        if derived.get(field) and text in str(derived[field]):
            sev = "error" if block.get("unmapped", True) else "warning"
            add(sev, field, f"{text}"
                            + (" (this person is terminated - dropping them "
                               "from the load fixes it too)" if terminated else ""))

    if terminated and not derived["term_date"]:
        add("warning", "term_date",
            "Inactive in Concur but ADP has no Termination Date.")

    if (emp.get("duplicate_rows") or 1) > 1:
        add("warning", "duplicate_rows",
            _s(emp.get("duplicate_note")) or
            "ADP sent more than one row for this File Number.")

    if (_s(emp.get("roster")) != "non_us"
            and _s(emp.get("legal_country_code")) not in ("USA", "US", "")):
        # Only worth saying on the ADP roster, where a non-US country code
        # means somebody has turned up in the wrong load. On the non-US roster
        # it is the entire premise.
        add("warning", "legal_country_code",
            f"Non-US employee ({_s(emp.get('legal_country_code'))}) on the ADP "
            "roster - they may belong on the non-US tab instead.")

    # Org Unit 2 has to be a code from a Concur connected list, not a name.
    #
    # Concur's Run-20 result said so in as many words: "Field [OrgUnit2] is
    # configured as a member of connected list [*Division - Department (List
    # Depth: 2)], but it contains an invalid list code [Sales]" - and rejected
    # the record. The US roster derives this from the Org Map and sends digits,
    # which Concur accepts; the non-US tabs have it typed in by hand as
    # 'Production', 'SALES', 'G&A', 'Selling' and so on, and every one of those
    # will be refused.
    #
    # The test is the shape rather than a list, because the list lives in
    # Concur and we do not hold it - so this is a warning by default, not a
    # block. Set block_on.org_unit_2_code to keep them out of the file once
    # somebody confirms the codes; the message says which employees to ask
    # about. A blank Org Unit 2 is fine - Concur took three of those in the
    # same run.
    unit2 = _s(derived.get("org_unit_2"))
    if unit2 and not unit2.isdigit() and "Not Mapped" not in unit2:
        add("error" if block.get("org_unit_2_code") else "warning",
            "org_unit_2",
            f"Org Unit 2 is '{unit2}', which is a name rather than a code. "
            "Concur has this field on the connected list '*Division - "
            "Department' and rejects a value that is not in it - it refused "
            "'Sales' outright on run 20. Every code it has accepted so far is "
            "numeric. Blank is accepted; a department name is not.")

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
        # Ctry Code. The one column here that the US 305 tab does NOT have a
        # formula for - it is blank on the template, so the extract wrote it
        # blank, and Concur rejected 13 records outright with "Missing required
        # field - ctry_code". They were the 13 people who do not exist in
        # Concur yet; UPDATE does not re-require it for the other 119, which is
        # why a gap this size stayed invisible for eighteen runs.
        #
        # It is filled from the same derived value the template already trusts:
        # legal_country is AH on sheet 1, and the Locale Code in column I is
        # literally the language code with AH glued on the end. So 'en_US' in I
        # and an empty J were always contradicting each other. The non-US tab
        # types the country in by hand and agrees.
        "J":  ("field", "legal_country"),            # Ctry Code
        "L":  ("field", "org_unit_1"),               # Ledger Code
        "M":  ("field", "reimbursement_currency"),
        "O":  ("field", "concur_status"),            # Active (Y/N)
        "P":  ("field", "org_unit_1"),
        "Q":  ("field", "org_unit_2"),
        "W":  ("field", "pay_grade_code"),           # Custom 2 Salary Code
        # Custom 3 Expense Profile. The Salary Map has two columns - C
        # 'Expense Map' (Default / Grade 20 / Officers) and D 'Travel Map'
        # (General / Senior Leadership / VIP) - and the two 305 tabs read
        # different ones: the US tab has ='1'!AG5, which is VLOOKUP(...,4), the
        # Travel column; the non-US tab has VLOOKUP(...,3), the Expense column.
        # A column headed 'Expense Profie' fed from the Travel map is a
        # transposition in the workbook, and Concur has since said so - it
        # warned 120 times that 'General' is not a Custom 3 list item.
        # So both rosters now read the Expense column, and the roster override
        # this used to need is gone.
        "X":  ("field", "concur_profile"),           # Custom 3 Expense Profile
        "Z":  ("field", "org_unit_1"),               # Custom 5 Department
        "AB": ("field", "term_date"),                # Custom 7 Term Date
        "AC": ("field", "preferred_name"),           # Custom 8 Preferred Name
        "AP": ("field", "org_unit_1"),               # Custom 21 Expense Group
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
    # Update ID Information Import - SAP's "UpdateIDInformationImporter". SAP
    # is explicit that this, not the 305, is the record that should carry
    # Login ID changes: "the administrator is strongly encouraged to use this
    # record type for this purpose instead of any other record type." SAP
    # also requires it be run on its own, a day ahead of the 305 - never in
    # the same file (see ADP_Concur_Export.ADP_Concur_export_320).
    # "C" New Employee ID is left out of the map - there is no "pending
    # rename" value held anywhere for it, so like any other column this
    # template does not point at, it writes blank.
    "320": {
        "A": ("const", "320"),
        "B": ("field", "file_number"),               # Current Employee ID
        "D": ("field", "login_id"),                  # New Login ID
    },
}

# Where a roster's own tab points a column somewhere else.
#
# Empty, and worth saying why: the one entry this ever held was Custom 3, where
# the two 305 tabs read different Salary Map columns. That was reproduced
# rather than reconciled while which one was right was still a question for
# SAP. Concur's Run-18 result answered it - the US tab's value was rejected and
# the non-US tab's was not - so the two are reconciled in FIELD_MAP above and
# the override is no longer needed. The mechanism stays because the tabs have
# disagreed before and it is what let each extract match the tab it came from
# while they did.
FIELD_MAP_BY_ROSTER: dict[str, dict[str, tuple[str, str]]] = {}


# Used when the workbook layouts have not been captured, so the record still
# comes out the right width. Taken from the template Kelly is working from.
# The 100 record is not in the workbook at all - its width is SAP's.
DEFAULT_WIDTHS = {"100": 7, "305": 137, "350": 67, "360": 35, "320": 9}

# The 100 record's seven fields, in SAP's order. Every one is required, so the
# record is written in full rather than padded like the employee records.
IMPORT_SETTINGS_FIELDS = [
    "error_threshold",
    "password_generation",
    "existing_record_handling",
    "language_code",
    "validate_expense_group",
    "validate_payment_group",
]

PASSWORD_GENERATION = {"EMPID", "LOGINID", "TEXT", "SSO"}
EXISTING_RECORD_HANDLING = {"REPLACE", "UPDATE", "WARN", "IGNORE"}


def build_import_settings(cfg: dict) -> list[str]:
    """
    The 100 record: '100' then the six settings, exactly as SAP orders them.

    One per file, first line. It is not built from the workbook - the workbook
    has no 100 tab - so it comes from the config, and the two values with a
    fixed vocabulary are checked here rather than at the far end of a load.
    """
    settings = cfg.get("import_settings") or {}
    values = ["100"] + [_s(settings.get(f)) for f in IMPORT_SETTINGS_FIELDS]

    generation = values[2].upper()
    if generation not in PASSWORD_GENERATION:
        raise ValueError(
            f"Password Generation is '{values[2]}'; SAP accepts "
            + ", ".join(sorted(PASSWORD_GENERATION)) + ".")
    handling = values[3].upper()
    if handling not in EXISTING_RECORD_HANDLING:
        raise ValueError(
            f"Existing Record Handling is '{values[3]}'; SAP accepts "
            + ", ".join(sorted(EXISTING_RECORD_HANDLING)) + ".")
    for i, label in ((5, "Validate Expense Group"), (6, "Validate Payment Group")):
        if values[i].upper() not in ("Y", "N"):
            raise ValueError(f"{label} is '{values[i]}'; it must be Y or N.")
    if not values[4]:
        raise ValueError("Language Code is required on the 100 record.")

    values[2], values[3] = generation, handling
    values[5], values[6] = values[5].upper(), values[6].upper()
    return values


def layout_width(conn: sqlite3.Connection, record_type: str) -> int:
    """How many fields a record of this type carries."""
    row = conn.execute(
        "SELECT MAX(position) FROM ADP_Concur_Layouts WHERE record_type = ?",
        (record_type,)).fetchone()
    return (row and row[0]) or DEFAULT_WIDTHS.get(record_type, 0)


def build_record(emp: dict, record_type: str, width: int, cfg: dict) -> list[str]:
    """One record as a list of field values, padded to the layout width."""
    out = [""] * width
    fields = dict(FIELD_MAP[record_type])
    fields.update((FIELD_MAP_BY_ROSTER.get(_s(emp.get("roster")), {})
                   .get(record_type, {})))
    for ref, (kind, value) in fields.items():
        idx = column_ref_to_index(ref) - 1
        if idx >= width:
            continue
        if kind == "const":
            out[idx] = value
        elif kind == "config":
            # The non-US tab carries a password per person; the US roster has
            # none and takes the configured one.
            if value == "password":
                out[idx] = _s(emp.get("password")) or _s(cfg.get("password", ""))
            else:
                out[idx] = _s(cfg.get(value, ""))
        else:
            out[idx] = _s(emp.get(value))

    # Blanked last, so a column is emptied whether it came from the base map,
    # a roster override or a constant. Case and stray spaces are forgiven
    # because these are typed into a config file by hand.
    for ref in (cfg.get("blank_fields") or {}).get(record_type, []) or []:
        ref = _s(ref).upper()
        if not ref:
            continue
        idx = column_ref_to_index(ref) - 1
        if 0 <= idx < width:
            out[idx] = ""
    return out


def selected_employees(conn: sqlite3.Connection, record_type: str,
                       cfg: dict, keys: list[int] | None = None,
                       roster: str = "") -> list[dict]:
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
    if roster:
        sql += " AND e.roster = ?"
        args.append(roster)
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
