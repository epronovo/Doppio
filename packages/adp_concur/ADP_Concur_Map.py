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

from ADP_Concur_Db import DERIVED_COLUMNS, SOURCES

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
    "invoice_limit": "Invoice Approval Not Mapped",
    # UKG's sheet words its own country miss differently from ADP's, and the
    # wording is the point - a value that fails should be recognisable to
    # whoever is looking at the tab it came from.
    "ukg_country": "Country Not Mapped",
    # The retired non-US tab's own wording, kept so an older database's stored
    # values still read as failures rather than as data.
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
    # Per-source overrides, merged over login_id above. Empty because both
    # sources currently agree; kept because they have disagreed before and the
    # mechanism is what let the extract match each tab while they did.
    "login_id_by_source": {},
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
    # Which record types each source writes. In config rather than in code
    # because it is a real question and it has already changed twice: the July
    # file carried 305 and 360 only, the September workbook added 350 and 700,
    # and every 350 Concur has seen so far has been rejected on its Travel
    # Class Name. Drop "350" from the adp list to go back to what loaded.
    "records_by_source": {
        "adp": ["305", "350", "360", "700"],
        "ukg": ["305", "360", "700"],
        "manual": ["305", "350", "360", "700"],
    },
    # Which people reach each record type. Always 'all' - active and
    # terminated people alike - except: 'invoice' means invoice_access = 'Y',
    # which is what the 700 is for; 'off' leaves the record type out of the
    # extract entirely, for the days every 350 Concur has seen gets rejected
    # on its Travel Class Name and the simplest fix is to stop sending it.
    "scope": {"305": "all", "350": "all", "360": "all", "700": "invoice"},
    # Columns to write empty whatever the field map says, per record type.
    #
    # This exists because a column can be right by the workbook and still wrong
    # by the tenant. Concur's Run-18 result warned 120 times that Custom 5 -
    # the Org Unit 2 department code - "could not be resolved to an existing
    # custom list item", and the July file that loaded cleanly left Custom 5
    # blank on all 216 of its 305 records. Whether the list is missing from the
    # tenant or the column is simply not wanted is a question for SAP, and this
    # is how you answer it without editing code: add "Z" to blank it, take it
    # out to send it again.
    #
    # Nothing is blanked by default - the extract still reproduces the
    # workbook, and every departure from it stays a decision somebody made.
    "blank_fields": {"305": [], "350": [], "360": [], "700": []},
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
    # The Salary Map grew a third column, 'Invoice Approval' - the currency
    # amount a grade may approve, which becomes the 700 record's limit. The
    # key is trimmed because the UKG sheet looks it up with TRIM() and its
    # Salary Grade arrives padded.
    maps["salary"] = {
        _s(r["pay_grade_code"]): (_s(r["expense_map"]), _s(r["travel_map"]),
                                  _s(r["invoice_map"]))
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

    # The country blocks beside the Country and Language maps: country ->
    # locale, and country -> currency code.
    maps["locale"] = {
        _s(r["country_code"]).upper(): _s(r["locale_code"])
        for r in conn.execute("SELECT * FROM ADP_Concur_LocaleMap")
    }
    maps["country_ref"] = {
        _s(r["country_code"]).upper(): _s(r["currency_code"])
        for r in conn.execute("SELECT * FROM ADP_Concur_CountryRef")
    }

    # The Country Map now carries the currency as well as the two-character
    # code, which is where UKG gets its reimbursement currency - it has no
    # business unit, so the Org Map route the ADP sheet uses is not open to it.
    maps["country_currency"] = {
        _s(r["adp_country"]).upper(): _s(r["currency_code"])
        for r in conn.execute("SELECT * FROM ADP_Concur_CountryMap")
    }

    # Invoice access by name. Keyed on File Number, and the workbook trims it
    # ('203011   ' appears with trailing spaces), so it is trimmed here too.
    maps["invoice_access"] = {
        _s(r["file_number"]): _s(r["access"]).upper()
        for r in conn.execute("SELECT * FROM ADP_Concur_InvoiceMap")
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
    rule.update((cfg.get("login_id_by_source") or {}).get(
        _s(emp.get("source")) or "adp", {}))
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


DERIVED_FIELDS = [name for _, name in DERIVED_COLUMNS]


def ADP_Concur_invoice_limit(emp: dict, maps: dict, access: str) -> str:
    """
    ADP AH: =IFERROR(VLOOKUP(AA5,'Salary Map'!A:E,5,FALSE),"Invoice Approval Not Mapped")
    UKG BE: =IF(BL2="N",0,IFERROR(VLOOKUP(TRIM(Q2),'Salary Map'!A:E,5,FALSE),"Salary Not Mapped"))

    How much this person may approve, from their pay grade. It is the whole
    content of the 700 record.

    The two sheets differ in one way that matters and is kept: UKG short-
    circuits to 0 when invoice access is N, ADP does not. So an ADP employee
    with no access still carries a limit in the sheet, and it simply never
    reaches a 700 because the 700 is conditional on access anyway. Reproduced
    rather than reconciled - it changes nothing in the file, and the day
    somebody reads the column it should say what their sheet says.

    The lookup key is trimmed for both, because UKG wraps it in TRIM() and its
    Salary Grade does arrive padded.
    """
    if _s(emp.get("source")) == "ukg" and access == "N":
        return "0"
    hit = maps["salary"].get(_s(emp.get("pay_grade_code")))
    if not hit:
        return UNMAPPED["invoice_limit"]
    return hit[2]


def ADP_Concur_invoice_access(emp: dict, maps: dict, limit_hit) -> str:
    """
    ADP AO: =IFERROR(VLOOKUP(P5,'Invoice Exception Map'!A:C,3,FALSE),IF(AH5>0,"Y","N"))
    UKG BL: =IFERROR(VLOOKUP(D2,'Invoice Exception Map'!A:C,3,FALSE),"N")

    Whether this person gets Concur Invoice at all. It drives four flags on the
    360, both 360 approver fields, and whether a 700 is written.

    The Invoice Exception Map is checked first for both, and it is the whole
    answer for UKG: nobody in UKG gets invoice access by grade, only by name.
    ADP falls back to the pay grade - an approval limit above zero means
    access. That asymmetry is the sheets', not mine, and it is why a UKG
    employee on the same grade as an ADP one can come out differently.
    """
    named = maps["invoice_access"].get(_s(emp.get("file_number")))
    if named in ("Y", "N"):
        return named
    if _s(emp.get("source")) == "ukg":
        return "N"
    # IF(AH5>0,"Y","N") - a text limit that is not a number is not > 0.
    try:
        return "Y" if float(str(limit_hit).replace(",", "")) > 0 else "N"
    except (TypeError, ValueError):
        return "N"


def derive_ukg(emp: dict, maps: dict, cfg: dict) -> dict:
    """
    The same fourteen values for somebody out of UKG.

    A separate function rather than branches inside the ADP one, because UKG
    is a different system describing the same people differently - not a
    variant of the ADP rules. Half of what ADP looks up, UKG already answers:
    the supervisor is already a bare employee number, the org units are codes
    on the record, and the status is already Y or N. What is left is four
    lookups, and each quotes the UKG sheet's own formula.

    Three of the fourteen are deliberately empty because UKG's derived block
    leaves them empty: BG Local Code, BI Preferred Name and BK Term Date have
    no formula at all. The 305 UKG tab computes its own locale from the
    country rather than reading BG, which is why locale_code is filled here
    and the others are not.
    """
    country = _s(emp.get("legal_country_code")).upper()

    # BF: =IFERROR(VLOOKUP(AM2,'Country Map'!A:E,2,FALSE),"Country Not Mapped")
    legal_country = maps["country"].get(_s(emp.get("legal_country_code"))) \
        or UNMAPPED["ukg_country"]

    # 305 UKG column I: =IFERROR(VLOOKUP(J2,'Language Map'!M:O,3,FALSE),"en_"&J2)
    # J is the Ctry Code, so the locale is looked up from the *derived* country
    # rather than from the raw one - 'US', not 'USA'.
    locale = (maps["locale"].get(legal_country.upper())
              or (f"en_{legal_country}" if legal_country
                  and "Not Mapped" not in legal_country else ""))

    # BH: =IFERROR(VLOOKUP(AM2,'Country Map'!A:E,4,FALSE),"Country Not Mapped")
    currency = maps["country_currency"].get(country) or UNMAPPED["ukg_country"]

    # BC/BD: =IFERROR(VLOOKUP(TRIM(Q2),'Salary Map'!A:E,3|4,FALSE),"Salary Code Not Mapped")
    grade = maps["salary"].get(_s(emp.get("pay_grade_code")))

    # BJ: =Y2 - the Employment Status Code, already A/Y/N shaped. UKG writes
    # 'A' for active, which the Status Map turns into Y like any other value.
    status = ADP_Concur_status(emp, maps)

    limit_hit = grade[2] if grade else UNMAPPED["invoice_limit"]
    access = ADP_Concur_invoice_access(emp, maps, limit_hit)

    return {
        # AZ: =R2. Already a bare employee number - no payroll prefix to strip.
        # The Supervisor Map still wins, for the same reason it does on ADP.
        "supervisor_id": (maps["supervisor"].get(_s(emp.get("file_number")))
                          if _s(emp.get("file_number")) in maps["supervisor"]
                          else _s(emp.get("supervisor_id_raw"))),
        # BA: =U2, the Site Location Code.  BB: =AI2, Org Level 3 Code.
        "org_unit_1": _s(emp.get("site_location_code")),
        "org_unit_2": _s(emp.get("org_level_3_code")),
        "concur_profile": grade[0] if grade else UNMAPPED["concur_profile"],
        "travel_profile": grade[1] if grade else UNMAPPED["travel_profile"],
        "invoice_limit": ADP_Concur_invoice_limit(emp, maps, access),
        "legal_country": legal_country,
        "locale_code": locale,
        "reimbursement_currency": currency,
        # BI and BK have no formula on the UKG sheet.
        "preferred_name": "",
        "concur_status": status,
        "term_date": "",
        "invoice_access": access,
        "login_id": ADP_Concur_login_id(emp, cfg),
    }


def derive_adp(emp: dict, maps: dict, cfg: dict) -> dict:
    """Every derived value for an ADP employee, in the workbook's own order."""
    legal_country = ADP_Concur_legal_country(emp, maps)
    concur_status = ADP_Concur_status(emp, maps)
    grade = maps["salary"].get(_s(emp.get("pay_grade_code")))
    limit_hit = grade[2] if grade else UNMAPPED["invoice_limit"]
    access = ADP_Concur_invoice_access(emp, maps, limit_hit)
    return {
        "supervisor_id": ADP_Concur_supervisor_id(emp, maps),
        "org_unit_1": ADP_Concur_org_unit_1(emp, maps),
        "org_unit_2": ADP_Concur_org_unit_2(emp, maps),
        "concur_profile": ADP_Concur_concur_profile(emp, maps),
        "travel_profile": ADP_Concur_travel_profile(emp, maps),
        "invoice_limit": ADP_Concur_invoice_limit(emp, maps, access),
        "legal_country": legal_country,
        "locale_code": ADP_Concur_locale_code(emp, maps, legal_country),
        "reimbursement_currency": ADP_Concur_reimbursement_currency(emp, maps),
        "preferred_name": ADP_Concur_preferred_name(emp),
        "concur_status": concur_status,
        "term_date": ADP_Concur_term_date(emp, concur_status),
        "invoice_access": access,
        "login_id": ADP_Concur_login_id(emp, cfg),
    }


def derive_one(emp: dict, maps: dict, cfg: dict) -> dict:
    """
    Every derived value for one employee, by whichever system sent them.

    Somebody keyed in by hand takes the ADP rules: they are typed into the
    app in ADP's terms, so ADP's lookups are the ones that apply to them.
    """
    if _s(emp.get("source")) == "ukg":
        return derive_ukg(emp, maps, cfg)
    return derive_adp(emp, maps, cfg)


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
    exceptions.extend(check_passwords(conn))

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


def check_passwords(conn: sqlite3.Connection) -> list[tuple]:
    """
    Passwords that look like a spreadsheet fill series rather than a decision.

    Dragging 'Welcome01' down a column gives Welcome02, Welcome03 and so on,
    which is what the non-US tab carries: 82 people, 82 different passwords,
    numbered 1 to 82 with no gaps. That is Excel's autofill, not a password
    policy, and it is worth saying out loud before it becomes 82 real
    credentials - especially as the 100 record's TEXT setting is what makes
    Concur use them.
    """
    rows = [dict(r) for r in conn.execute(
        "SELECT employee_key, file_number, "
        "legal_last_name || ', ' || legal_first_name AS name, password, source "
        "FROM ADP_Concur_Employees "
        "WHERE row_state <> 'deleted' AND password <> '' ORDER BY file_number")]
    series: dict[str, list] = {}
    for r in rows:
        m = re.fullmatch(r"(.*?)(\d+)", _s(r["password"]))
        if m:
            series.setdefault(m.group(1), []).append((int(m.group(2)), r))

    out = []
    for stem, entries in series.items():
        numbers = sorted(n for n, _ in entries)
        # A run is only suspicious when it is long and unbroken - two people
        # who happen to be Welcome01 and Welcome02 prove nothing.
        if len(numbers) < 5 or numbers != list(range(numbers[0], numbers[0] + len(numbers))):
            continue
        for _n, r in entries:
            out.append((r["employee_key"], r["file_number"], r["name"], "warning",
                        "password",
                        f"These passwords run {stem}{numbers[0]:02d} "
                        f"to {stem}{numbers[-1]:02d} with no gaps - a spreadsheet "
                        "fill series rather than {len} chosen passwords. The 100 "
                        "record's TEXT setting is what sends them to Concur."
                        .replace("{len}", str(len(numbers)))))
    return out


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

    # Approvers who report across the ADP/UKG line. This used to be an error
    # about load order, because the two rosters were two files and the file
    # holding the approver had to go first. One combined file, sorted so that
    # an approver is always written above the people pointing at them, is what
    # retired that problem - so this is now a note that the chain crosses
    # systems, which is worth seeing when a name looks wrong, and nothing more.
    for r in conn.execute(
        """
        SELECT e.employee_key, e.file_number,
               e.legal_last_name || ', ' || e.legal_first_name AS name,
               e.source, e.supervisor_id, s.source AS sup_source,
               s.legal_last_name || ', ' || s.legal_first_name AS sup_name
          FROM ADP_Concur_Employees e
          JOIN ADP_Concur_Employees s ON s.file_number = e.supervisor_id
         WHERE e.row_state <> 'deleted' AND s.row_state <> 'deleted'
           AND e.supervisor_id <> '' AND e.source <> s.source
           AND e.source IN ('adp','ukg') AND s.source IN ('adp','ukg')
        """
    ):
        out.append((r["employee_key"], r["file_number"], r["name"], "warning",
                    "supervisor_chain",
                    f"Reports to {r['sup_name']} ({r['supervisor_id']}), who "
                    f"comes from {r['sup_source'].upper()} where this record "
                    f"comes from {r['source'].upper()}. Both are in the same "
                    "file and the approver is written first, so this is only "
                    "worth knowing, not fixing."))

    # Approvers who have left. Worth its own check now that the file is sorted
    # active-first: everybody active is written before anybody inactive, so an
    # active person reporting to a leaver is the one pairing the sort cannot
    # put in order, and their approver lands below them. Concur will not
    # resolve a terminated approver anyway, so the ordering is the symptom
    # rather than the disease - but this is where it becomes visible.
    for r in conn.execute(
        """
        SELECT e.employee_key, e.file_number,
               e.legal_last_name || ', ' || e.legal_first_name AS name,
               e.supervisor_id,
               s.legal_last_name || ', ' || s.legal_first_name AS sup_name
          FROM ADP_Concur_Employees e
          JOIN ADP_Concur_Employees s ON s.file_number = e.supervisor_id
         WHERE e.row_state <> 'deleted' AND s.row_state <> 'deleted'
           AND e.supervisor_id <> ''
           AND e.concur_status = 'Y' AND s.concur_status = 'N'
        """
    ):
        out.append((r["employee_key"], r["file_number"], r["name"], "warning",
                    "supervisor_chain",
                    f"Approver {r['sup_name']} ({r['supervisor_id']}) is "
                    "inactive. An active employee cannot be approved by "
                    "somebody who has left - they need a live approver, or "
                    "the leaver needs to stay active until the chain moves."))

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

    # Somebody in both exports at once. The two systems overlap - a person
    # moved from ADP to UKG can appear on both sheets in the same cut - and
    # since both merge on File Number into one row, whichever loaded last
    # silently won. Saying so is the only way that gets noticed.
    if _s(emp.get("duplicate_note")).startswith("Both "):
        add("warning", "source",
            _s(emp.get("duplicate_note")))

    # Invoice access granted by name, but with nothing to approve.
    #
    # The Invoice Exception Map wins over the pay grade, so somebody can be
    # given access while their grade's approval limit is 0 - and the 700 then
    # goes out saying they may approve up to nothing. The workbook does the
    # same thing, so this is faithful rather than wrong, but it is almost
    # certainly not what was meant: either the grade needs a limit or the
    # person does not need the access.
    if (_s(derived.get("invoice_access")) == "Y"
            and _s(derived.get("invoice_limit")) in ("0", "0.0", "")):
        add("warning", "invoice_limit",
            "Has invoice access from the Invoice Exception Map, but their pay "
            f"grade ({_s(emp.get('pay_grade_code')) or 'none'}) carries an "
            "approval limit of "
            f"{_s(derived.get('invoice_limit')) or 'nothing'}. The 700 will "
            "say they can approve up to zero.")

    # Org Unit 2 has to be a code from a Concur connected list, not a name.
    #
    # Concur's Run-20 result said so in as many words: "Field [OrgUnit2] is
    # configured as a member of connected list [*Division - Department (List
    # Depth: 2)], but it contains an invalid list code [Sales]" - and rejected
    # the record. ADP derives this from the Org Map and sends digits,
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
        "J":  ("field", "legal_country"),            # Ctry Code
        "L":  ("field", "org_unit_1"),               # Ledger Code
        "M":  ("field", "reimbursement_currency"),
        "O":  ("field", "concur_status"),            # Active (Y/N)
        "P":  ("field", "org_unit_1"),
        "Q":  ("field", "org_unit_2"),
        "W":  ("field", "pay_grade_code"),           # Custom 2 Salary Code
        # Custom 3 Expense Profile. Both 305 tabs now read the Salary Map's
        # Expense column - the transposition that had the ADP tab reading the
        # Travel column is gone from the 15 September workbook, which is what
        # Concur's Run-20 result had already told us.
        "X":  ("field", "concur_profile"),
        # Custom 5 Department: =P5 on both tabs. It is Org Unit 1 now, not the
        # department code - the workbook changed it after Concur rejected the
        # department codes as invalid list items.
        "Z":  ("field", "org_unit_1"),
        "AB": ("field", "term_date"),                # Custom 7 Term Date
        "AC": ("field", "preferred_name"),           # Custom 8 Preferred Name
        "AP": ("field", "org_unit_1"),               # Custom 21 Expense Group
        "BG": ("field", "supervisor_id"),            # Expense Report Approver
        "BK": ("const", "Y"),                        # Expense User
        "BL": ("const", "Y"),                        # Expense / Cash Advance Approver
        # Invoice User and Invoice Approver are no longer flat Y. Both tabs
        # feed them from Invoice Access, so somebody without it is loaded
        # without Concur Invoice rather than given it and left unable to use it.
        "BU": ("field", "invoice_access"),           # Invoice User
        "BV": ("field", "invoice_access"),           # Invoice Approver
        "CE": ("const", "Y"),                        # Future Use 2
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
        "C":  ("field", "invoice_access"),           # Invoice User Role
        "D":  ("field", "invoice_access"),           # Invoice Approver Role
        "I":  ("field", "invoice_access"),           # Purchase Request User
        "J":  ("field", "invoice_access"),           # Purchase Request Approver
        "R":  ("field", "supervisor_id"),            # Default PR Approver
        "S":  ("field", "supervisor_id"),            # Payment Approver
        "AC": ("const", "Y"),                        # Display Image In-line
        "AD": ("const", "Y"),                        # Auto Open Image
    },
    # The 700: a payment request approval authority, one per person who has
    # invoice access. Everything on it is conditional on that - the tabs write
    # the whole row as =IF($AO5="Y", ..., "") - so the record is not built at
    # all for anybody else rather than being built empty. selected_employees()
    # applies that filter; see the 'invoice' scope.
    "700": {
        "A": ("const", "700"),
        "B": ("const", "REQ"),                       # Approval Type
        "C": ("field", "file_number"),
        "O": ("field", "invoice_limit"),             # Approval Limit
        "P": ("field", "reimbursement_currency"),    # Approval Limit Currency
    },
}

# Where a source's own tab points a column somewhere else.
#
# Four differences between the ADP and UKG tabs, all of them real:
#
#  * 305 C Middle Name - UKG has no middle name column at all.
#  * 305 AB / AC - UKG's derived block leaves Term Date and Preferred Name
#    empty, so its tab has no formula in either.
#  * 360 AC Display Image In-line - ADP hard-codes Y, UKG feeds it from
#    Invoice Access; and AD Auto Open Image is Y on ADP and N on UKG.
#  * 700 P Approval Limit Currency - ADP writes the employee's reimbursement
#    currency, UKG hard-codes USD, because a UKG approval limit is stated in
#    dollars whatever the employee is paid in.
#
# Each is written as the tab writes it. A source with no entry here takes the
# base map above.
FIELD_MAP_BY_SOURCE: dict[str, dict[str, dict[str, tuple[str, str]]]] = {
    "ukg": {
        "305": {
            "C":  ("const", ""),                     # no middle name in UKG
            "AB": ("const", ""),                     # no Term Date
            "AC": ("const", ""),                     # no Preferred Name
        },
        "360": {
            "AC": ("field", "invoice_access"),       # Display Image In-line
            "AD": ("const", "N"),                    # Auto Open Image
        },
        "700": {
            "P": ("const", "USD"),                   # Approval Limit Currency
        },
    },
}

# Used when the workbook layouts have not been captured, so the record still
# comes out the right width. Taken from the template Kelly is working from.
# The 100 record is not in the workbook at all - its width is SAP's.
DEFAULT_WIDTHS = {"100": 7, "305": 137, "350": 67, "360": 35, "700": 16}

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
    fields.update((FIELD_MAP_BY_SOURCE.get(_s(emp.get("source")), {})
                   .get(record_type, {})))
    for ref, (kind, value) in fields.items():
        idx = column_ref_to_index(ref) - 1
        if idx >= width:
            continue
        if kind == "const":
            out[idx] = value
        elif kind == "config":
            # A hand-keyed person can carry their own password; everybody
            # else takes the configured one.
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



def _sources_writing(record_type: str) -> list[str]:
    """Which sources carry a tab for this record type."""
    return [name for name, spec in SOURCES.items()
            if record_type in spec["records"]]


def selected_employees(conn: sqlite3.Connection, record_type: str,
                       cfg: dict, keys: list[int] | None = None,
                       source: str = "") -> list[dict]:
    """
    The people who belong in one record type.

    Four things decide it: the per-record include flag on the employee (so a
    single person can be held back without touching anything else), the
    configured scope - always 'all' except the 700's 'invoice', or 'off' to
    leave the record type out entirely - the exception list, which keeps
    blocking errors out of the file, and `keys`, which narrows the whole
    thing to a chosen set of people.

    `keys` is how a pilot load is built: pick one manager's organisation on the
    Hierarchy tab and the extract carries those people and nobody else. An
    empty list is not the same as None - it means nothing was chosen, and it
    produces an empty file rather than the whole company.
    """
    scope = (cfg.get("scope") or {}).get(record_type, "all")
    if scope == "off":
        return []
    sql = (f"SELECT e.* FROM ADP_Concur_Employees e "
           f"WHERE e.row_state <> 'deleted' AND e.include_{record_type} = 1")
    args: list = []
    if source:
        sql += " AND e.source = ?"
        args.append(source)
    # A source only writes the record types its own tabs carry: UKG has no 350.
    sql += (" AND e.source IN (" + ",".join(
        "?" for _ in _sources_writing(record_type)) + ")")
    args += _sources_writing(record_type)
    if scope == "invoice":
        # The 700 exists only for people with invoice access - the tabs write
        # the whole row as =IF(AccessIsY, ..., "") and an empty row is not a
        # record. Filtering here rather than emitting blanks keeps the file
        # free of 700s that say nothing.
        sql += " AND e.invoice_access = 'Y'"
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
