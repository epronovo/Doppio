import os, sqlite3, sys, json
from datetime import date
sys.path.insert(0, "/Users/ericpronovost/Doppio/packages/adp_concur")
import ADP_Concur_Map as M
import ADP_Concur_Db as D
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "/Users/ericpronovost/Doppio/packages/adp_concur/ADP_Concur_Mapping_Guide.xlsx"
conn = sqlite3.connect("file:" + os.path.expanduser("~/sqlite/doppio.db") + "?mode=ro", uri=True)
conn.row_factory = sqlite3.Row
cfg = M.load_config()

F = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(name=F, bold=True, color="FFFFFF", size=10)
BODY = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
TITLE = Font(name=F, size=14, bold=True, color="1F3864")
SUB = Font(name=F, size=10, italic=True, color="595959")
GREY = PatternFill("solid", fgColor="F2F2F2")
DIFF = PatternFill("solid", fgColor="FFF2CC")   # ADP and UKG differ
CONST = PatternFill("solid", fgColor="E2EFDA")  # constant value
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")

wb = Workbook()
wb.remove(wb.active)


def sheet(name, title, subtitle, headers, rows, widths, fills=None):
    ws = wb.create_sheet(name)
    ws["A1"] = title; ws["A1"].font = TITLE
    ws["A2"] = subtitle; ws["A2"].font = SUB
    hr = 4
    for i, h in enumerate(headers, 1):
        c = ws.cell(hr, i, h); c.font = HDR_FONT; c.fill = HDR_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BORDER
    for r, row in enumerate(rows, hr + 1):
        fill = fills[r - hr - 1] if fills else None
        for i, v in enumerate(row, 1):
            c = ws.cell(r, i, v); c.font = BODY; c.alignment = WRAP; c.border = BORDER
            if fill: c.fill = fill
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(hr + 1, 1)
    if rows:
        ws.auto_filter.ref = f"A{hr}:{get_column_letter(len(headers))}{hr + len(rows)}"
    ws.sheet_view.showGridLines = False
    return ws


# ------------------------------------------------------------ lookups
ADP_LETTERS = "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z AA AB".split()
adp_src = {c: (h, ADP_LETTERS[i]) for i, (h, c) in enumerate(D.ADP_COLUMNS)}
ukg_src = {}
for h, c in D.UKG_COLUMNS:
    ukg_src[c] = h
UKG_LETTER = {}
import re, inspect
for line in inspect.getsource(D).splitlines():
    m = re.match(r'\s*\("([^"]+)", "([a-z0-9_]+)"\),\s*#\s*([A-Z]+)\s*$', line)
    if m and m.group(2) in ukg_src and m.group(1) == ukg_src[m.group(2)]:
        UKG_LETTER[m.group(2)] = m.group(3)
derived_label = {c: h for h, c in D.DERIVED_COLUMNS}
extra_raw = {"password": "Password (hand-keyed employees only)",
             "employee_name_raw": "Employee Name"}


def describe(kind, value, source):
    if kind == "const":
        return f'Constant "{value}"' if value != "" else "(blank - no source)"
    if kind == "config":
        return (f'Config "password" (currently "{cfg.get("password")}"); '
                "a hand-keyed employee's own password wins")
    if value in derived_label:
        return f"Derived: {derived_label[value]}  [{value}]"
    if source == "ukg":
        if value in ("legal_first_name", "legal_last_name"):
            part = "text after" if value == "legal_first_name" else "text before"
            return (f"UKG: Employee Name (Last Suffix, First MI) (col F), {part} the comma  [{value}]")
        if value in ukg_src:
            return f"UKG: {ukg_src[value]} (col {UKG_LETTER.get(value, '?')})  [{value}]"
        return f"(UKG has no source - blank)  [{value}]"
    if value in adp_src:
        h, l = adp_src[value]
        return f"ADP: {h} (col {l})  [{value}]"
    return f"{extra_raw.get(value, value)}  [{value}]"


def field_for(rt, ref, source):
    fm = dict(M.FIELD_MAP.get(rt, {}))
    fm.update(M.FIELD_MAP_BY_SOURCE.get(source, {}).get(rt, {}))
    cc = dict(M.CONDITIONAL_COPIES.get(rt, {}))
    cc.update(M.CONDITIONAL_COPIES_BY_SOURCE.get(source, {}).get(rt, {}))
    if ref in (cfg.get("blank_fields") or {}).get(rt, []):
        return "Blanked by config blank_fields"
    if ref in cc:
        flag, src = cc[ref]
        what = f'constant "{src[1]}"' if isinstance(src, tuple) else f"column {src}"
        return f"If column {flag} = Y then {what}, else blank"
    if ref in fm:
        return describe(*fm[ref], source)
    return ""


# ------------------------------------------------------------ 1 overview
scope = cfg.get("scope", {})
ov = wb.create_sheet("Overview")
ov.sheet_view.showGridLines = False
ov.column_dimensions["A"].width = 26; ov.column_dimensions["B"].width = 110
lines = [
    ("ADP / UKG -> SAP Concur Employee Import: Mapping Guide", TITLE),
    (f"Generated {date.today():%Y-%m-%d} from ADP_Concur_Map.py, ADP_Concur_Db.py and ADP_Concur_Config.json. "
     "Map-table tabs are a snapshot of the live database (~/sqlite/doppio.db).", SUB),
    None,
    ("How the data flows", BOLD),
    ("1. Import", "The workbook is loaded (ADP_Concur_Import.py). The ADP export (sheet '1' / 'ADP') and the UKG export are merged "
     "on File Number / Employee Number into ADP_Concur_Employees; 'source' says which system sent each row. The map tabs load "
     "into their own ADP_Concur_*Map tables, and the 305/350/360/700 template tabs load into ADP_Concur_Layouts."),
    ("2. Derive", "ADP_Concur_derive() recomputes the derived columns (the workbook's AC..AO for ADP, AZ..BL for UKG) from "
     "the map tables. It runs after every import or map edit, and it rebuilds the exception list."),
    ("3. Build records", "build_record() fills each Concur record position from FIELD_MAP, then applies the per-source "
     "overrides, the conditional copies and the config blank_fields. A position with no mapping is written empty."),
    ("4. Extract", "One comma-delimited file: a 100 (Import Settings) line first, then the 305s, 350s, 360s and 700s "
     "('by_type' order). The 305s are sorted so an approver comes before the people who report to them, and the other "
     "record types follow that order. Employees with any 'error' exception are left out."),
    None,
    ("Sources", BOLD),
    ("ADP", "Uses the ADP rules. Record types: " + ", ".join(D.SOURCES["adp"]["records"])),
    ("UKG", "Uses the UKG rules (derive_ukg). Record types: " + ", ".join(D.SOURCES["ukg"]["records"]) + " (UKG has no 350)."),
    ("Added by hand", "Keyed into the app. Uses the ADP rules and can carry its own password."),
    None,
    ("Current scope (config)", BOLD),
] + [(f"Record {rt}", {"all": "Everyone included (active and terminated)",
                        "invoice": "Only employees with Invoice Access = Y",
                        "off": "Not written - left out of the extract"}.get(v, v))
     for rt, v in scope.items()] + [
    None,
    ("Tabs in this guide", BOLD),
    ("Source Columns", "Every ADP and UKG export column and the database field it loads into."),
    ("Derived Fields", "Each calculated value: the rule for ADP and for UKG, the map table it uses, and what it returns when the lookup finds nothing."),
    ("Record 100", "The Import Settings line."),
    ("Record 305/350/360/700", "Every position of each Concur record and what fills it, for ADP and for UKG."),
    ("Map Tables", "What each lookup table is keyed on and what it returns."),
    ("Map - *", "The current contents of each lookup table."),
    ("Config", "The settings that change the output, with the current value and the code default."),
    ("Validation", "The checks that raise exceptions, and whether each one keeps the employee out of the file."),
    None,
    ("Colour legend", BOLD),
    ("Yellow row", "ADP and UKG fill this position differently."),
    ("Green row", "A constant, the same for every employee."),
    ("Grey row", "Not mapped: the position is written empty."),
    ("[name]", "The database column (ADP_Concur_Employees) the value is read from."),
]
r = 1
for item in lines:
    if item is None:
        r += 1; continue
    if isinstance(item[1], Font):
        ov.cell(r, 1, item[0]).font = item[1]
    else:
        ov.cell(r, 1, item[0]).font = BOLD
        c = ov.cell(r, 2, item[1]); c.font = BODY; c.alignment = WRAP
    for cc in (ov.cell(r, 1), ov.cell(r, 2)):
        cc.alignment = WRAP if cc.column == 2 else Alignment(vertical="top")
    r += 1
ov["A2"].alignment = Alignment(wrap_text=False)
fills = {"Yellow row": DIFF, "Green row": CONST, "Grey row": GREY}
for row in ov.iter_rows(min_col=1, max_col=1):
    if row[0].value in fills:
        row[0].fill = fills[row[0].value]

# ------------------------------------------------------------ 2 source columns
rows = []
allcols = [c for _, c in D.ADP_COLUMNS] + [c for _, c in D.UKG_ONLY_COLUMNS]
for c in allcols:
    a = adp_src.get(c); u = ukg_src.get(c)
    rows.append([c, a[1] if a else "", a[0] if a else "", UKG_LETTER.get(c, "") if u else "", u or ""])
rows.append(["password", "", "", "", "(Hand-keyed employees only - typed into the app)"])
notes_fill = [DIFF if (r_[2] and r_[4] and r_[2] != r_[4]) else None for r_ in rows]
ws = sheet("Source Columns", "Source Columns -> Database",
           "Where each export column lands in ADP_Concur_Employees. Columns with the same meaning share one database field. "
           "Yellow = the same field under a different name in each system.",
           ["Database Field", "ADP Col", "ADP Export Heading", "UKG Col", "UKG Export Heading"],
           rows, [26, 9, 40, 9, 44], notes_fill)
n = len(rows) + 6
notes = [
    "UKG 'Employee Name (Last Suffix, First MI)' is split at the first comma into legal_last_name / legal_first_name (workbook: FIND/LEFT/MID).",
    "Legal / Preferred Address: Country Code (ADP) and Country Code (UKG) are both 3-character codes (USA, CHN); the Country Map turns them into Concur's 2-character code.",
    "Both exports merge on File Number. A person on both sheets ends up as one row, the last one loaded wins, and a warning is raised.",
]
ws.cell(n - 1, 1, "Notes").font = BOLD
for i, t in enumerate(notes):
    c = ws.cell(n + i, 1, t); c.font = BODY

# ------------------------------------------------------------ 3 derived fields
U = M.UNMAPPED
derived = [
 ("supervisor_id", "AC", "The Supervisor Map (keyed on the employee's File Number) wins. Otherwise characters 4-9 of the ADP Supervisor ID: 'KSY211756' -> '211756'.",
  "=IF(ISBLANK(H5),VLOOKUP(P5,'Supervisor Map'!A:D,4,FALSE),MID(H5,4,6))  (the app lets the map win over ADP as well)",
  "AZ", "The Supervisor Map wins. Otherwise the Supervisor Employee Number exactly as UKG sends it (=R2).",
  "Supervisor Map", "Blank. A warning if the person is active; blank on purpose means top of the hierarchy."),
 ("org_unit_1", "AD", "Org Map, looked up by Business Unit Description, returns Org Unit 1 (the first matching row wins).",
  "=IFERROR(VLOOKUP(T5,'Org Map'!A:E,4),\"BU Description Not Mapped\")",
  "BA / 305 P", "UKG Company Map, looked up by Site Location Code, returns Concur Company Code.",
  "Org Map (ADP); UKG Company Map (UKG)", f"ADP: \"{U['org_unit_1']}\"; UKG: \"{U['ukg_site_location']}\""),
 ("concur_expense_group", "-", "The same value as Org Unit 1.", "305 AP = ADP!AD",
  "305 AP", "UKG Company Map, looked up by Site Location Code, returns Expense Group Code.",
  "UKG Company Map", f"UKG: \"{U['ukg_site_location']}\""),
 ("concur_ledger_code", "-", "The same value as Org Unit 1.", "305 L = ADP!AD",
  "305 L", "UKG Company Map, looked up by Site Location Code, returns Ledger Code. (The workbook formula uses column index 7, which is past the end of the table. The app uses the Ledger Code column, as the heading intends.)",
  "UKG Company Map", f"UKG: \"{U['ukg_ledger_code']}\""),
 ("concur_custom_5", "-", "The same value as Org Unit 1.", "305 Z = ADP!AD",
  "305 Z", "UKG Company Map, looked up by Site Location Code, returns Custom 5 Code. (The workbook uses index 8; the app uses the Custom 5 Code column.)",
  "UKG Company Map", f"UKG: \"{U['ukg_custom_5']}\""),
 ("org_unit_2", "AE", "Org Map, looked up by Home Department Code, returns Org Unit 2. The app uses an exact match; the workbook formula does an approximate match.",
  "=IFERROR(VLOOKUP(V5,'Org Map'!C:E,3),\"Home Department Code Not Mapped\")",
  "BB", "Site Location Code 0077 or 177 (NGP Sweden) gives \"0000\". Otherwise Org Level 3 Code, and a blank becomes \"000\".",
  "Org Map (ADP only)", f"ADP: \"{U['org_unit_2']}\"; UKG: never unmapped"),
 ("concur_profile", "AF", "Salary Map, looked up by Pay Grade Code, returns the Expense Map value.",
  "=IFERROR(VLOOKUP(AA5,'Salary Map'!A:D,3,FALSE),\"Salary Code Not Mapped\")",
  "BC", "Salary Map, looked up by the trimmed Salary Grade, returns the Expense Map value.", "Salary Map", f"\"{U['concur_profile']}\""),
 ("travel_profile", "AG", "Salary Map, looked up by Pay Grade Code, returns the Travel Map value, used as it appears in the map (e.g. 'General US').",
  "=IFERROR(VLOOKUP(AA5,'Salary Map'!A:D,4,FALSE),\"Salary Code Not Mapped\")",
  "BD", "Same as ADP, using the trimmed Salary Grade.", "Salary Map", f"\"{U['travel_profile']}\""),
 ("invoice_limit", "AH", "Salary Map, looked up by Pay Grade Code, returns the Invoice Approval amount.",
  "=IFERROR(VLOOKUP(AA5,'Salary Map'!A:E,5,FALSE),\"Invoice Approval Not Mapped\")",
  "BE", "\"0\" if Invoice Access = N. Otherwise the Salary Map Invoice Approval amount for the trimmed Salary Grade.",
  "Salary Map", f"\"{U['invoice_limit']}\""),
 ("legal_country", "AI", "Country Map, looked up by the 3-character address country code, returns the 2-character Concur code.",
  "=IFERROR(VLOOKUP(Z5,'Country Map'!A:B,2,FALSE),\"Legal Country Not Mapped\")",
  "BF", "Country Map, looked up by Country Code, returns the 2-character code.", "Country Map",
  f"ADP: \"{U['legal_country']}\"; UKG: \"{U['ukg_country']}\""),
 ("locale_code", "AJ", "Language Map, looked up by Language Description, returns a stem such as 'en_'. If that finds nothing, the Org Map default language for the Business Unit is used. Legal Country is then added to the end: 'en_' + 'US' = 'en_US'.",
  "=IFERROR(VLOOKUP(X5,'Language Map'!A:C,3,FALSE),IFERROR(VLOOKUP(T5,'Org Map'!A:F,6,FALSE),\"ADP and BU Language Code Not Mapped\"))&AH5",
  "BG / 305 I", "Locale Map (Language Map columns M:O), looked up by the derived 2-character country. 'en_US' if the country is not in the map.",
  "Language Map, Org Map (ADP); Locale Map (UKG)", f"ADP: \"{U['locale_code']}\" + country; UKG: \"en_US\""),
 ("reimbursement_currency", "AK", "Org Map, looked up by Business Unit Description, returns Reimbursement Currency.",
  "=IFERROR(VLOOKUP(T5,'Org Map'!A:G,7),\"BU Currency Not Mapped\")",
  "BH", "Country Map, looked up by Country Code, returns Currency Code.", "Org Map (ADP); Country Map (UKG)",
  f"ADP: \"{U['reimbursement_currency']}\"; UKG: \"{U['ukg_country']}\""),
 ("preferred_name", "AL", "Preferred First Name + ' ' + Preferred Last Name. Blank if there is no preferred first name.",
  "=IF(ISBLANK(E5),\"\",TRIM(E5&\" \"&F5))", "BI", "Always blank (the UKG sheet has no formula here).", "-", "-"),
 ("concur_status", "AM", "Status Map, looked up by Position Status, returns Y (active) or N (inactive).",
  "=IFERROR(VLOOKUP(O5,'Status Map'!A:B,2,FALSE),\"Position Status Not Mapped\")",
  "BJ", "Status Map, looked up by Employment Status (UKG 'A' maps to Y).", "Status Map", f"\"{U['concur_status']}\""),
 ("term_date", "AN", "If Status = N: Termination Date formatted yyyymmdd. Otherwise blank.",
  "=IF(AL5=\"N\",TEXT(M5,\"yyyymmdd\"),\"\")", "BK", "Always blank (the UKG sheet has no formula here).", "-",
  "Blank. A warning if the person is inactive with no date."),
 ("invoice_access", "AO", "Invoice Exception Map (by File Number) returns Y or N and wins. Otherwise Y if the invoice limit is > 0, else N.",
  "=IFERROR(VLOOKUP(P5,'Invoice Exception Map'!A:C,3,FALSE),IF(AH5>0,\"Y\",\"N\"))",
  "BL", "Invoice Exception Map (by Employee Number). Otherwise N. UKG never gets access from the pay grade.",
  "Invoice Exception Map, Salary Map", "N"),
 ("login_id", "-", "Config login_id rule: the first non-blank of " + str(cfg["login_id"].get("sources"))
  + f", then prefix '{cfg['login_id'].get('prefix','')}', suffix '{cfg['login_id'].get('suffix','')}', replace_domain '{cfg['login_id'].get('replace_domain','')}', lowercase={cfg['login_id'].get('lowercase')}. "
  "With the current config this is the work email, unchanged.",
  "Not a workbook column (the workbook 305 builds it from the email)", "-", "Same rule as ADP.", "Config", "Blank, which is an error (no Login ID)."),
]
rows = [[f, derived_label.get(f, f), a, ar, af, u, ur, mp, miss] for f, a, ar, af, u, ur, mp, miss in derived]
sheet("Derived Fields", "Derived Fields",
      "Calculated by ADP_Concur_derive() and stored on ADP_Concur_Employees. ADP Col / UKG Col are the derived-block columns on each export sheet.",
      ["Database Field", "Label", "ADP Col", "ADP Rule", "Workbook Formula (ADP)", "UKG Col", "UKG Rule", "Map Table(s)", "Value When Not Found"],
      rows, [22, 20, 8, 48, 44, 10, 46, 22, 32], [DIFF if r_[3] != r_[6] and "Same" not in r_[6] else None for r_ in rows])

# ------------------------------------------------------------ 4 record 100
s = cfg["import_settings"]
meaning = {
 "error_threshold": "Errors allowed before the import stops. SAP says 0.",
 "password_generation": "EMPID | LOGINID | TEXT | SSO. SSO means the 305 password is ignored.",
 "existing_record_handling": "REPLACE | UPDATE | WARN | IGNORE. UPDATE writes only non-blank fields and never overwrites a password.",
 "language_code": "Language of any localised text in the file.",
 "validate_expense_group": "Y | N (SAP default Y).",
 "validate_payment_group": "Y | N (SAP default Y).",
}
rows = [[1, "A", "Record Type", "100", "Constant"]] + [
    [i + 2, get_column_letter(i + 2), k.replace("_", " ").title(), s.get(k), meaning[k]]
    for i, k in enumerate(M.IMPORT_SETTINGS_FIELDS)]
ws = sheet("Record 100", "Record 100 - Import Settings",
           "The first line of every file. Built from config import_settings, not from the workbook. Current line: "
           + ",".join(M.build_import_settings(cfg)),
           ["Pos", "Col", "Field", "Current Value", "Allowed Values / Meaning"], rows, [6, 6, 28, 16, 90])

# ------------------------------------------------------------ 5 record layouts
record_notes = {
 "305": "Employee profile. Every employee, active and terminated. Password (G) is ignored because Password Generation = "
        + s.get("password_generation") + ".",
 "350": "Travel profile. ADP and hand-keyed employees only; UKG has no 350. Current scope: " + scope.get("350", "all") + ".",
 "360": "Invoice / purchase request roles. Current scope: " + scope.get("360", "all") + ".",
 "700": "Payment approval authority. Written only for employees with Invoice Access = Y. Current scope: " + scope.get("700", "invoice") + ".",
}
for rt in ("305", "350", "360", "700"):
    layout = conn.execute("SELECT position, column_ref, heading, max_width FROM ADP_Concur_Layouts "
                          "WHERE record_type=? ORDER BY position", (rt,)).fetchall()
    rows, fl = [], []
    for L in layout:
        ref = L["column_ref"]
        heading = " ".join((L["heading"] or "").split())
        short = re.split(r"\s[•(*]|\s-\s|\s\*\*\*", heading)[0][:70]
        a = field_for(rt, ref, "adp")
        u = field_for(rt, ref, "ukg") if rt != "350" else "n/a - UKG writes no 350"
        rows.append([L["position"], ref, short, a, u, L["max_width"], heading if heading != short else ""])
        if not a and (rt == "350" or not u):
            fl.append(GREY)
        elif rt != "350" and a != u:
            fl.append(DIFF)
        elif a.startswith("Constant"):
            fl.append(CONST)
        else:
            fl.append(None)
    mapped = sum(1 for x in fl if x is not GREY)
    sheet(f"Record {rt}", f"Record {rt} - Field Mapping",
          record_notes[rt] + f" {len(rows)} positions, {mapped} filled. Every position is written, even when empty.",
          ["Pos", "Col", "Concur Field", "ADP Value", "UKG Value", "Max Len", "Full Template Heading"],
          rows, [6, 6, 34, 52, 52, 8, 60], fl)

# ------------------------------------------------------------ 6 map tables
maps = [
 ("Status", "ADP_Concur_StatusMap", "Status Map", "position_status", "concur_status", "Status (Y/N)"),
 ("Country", "ADP_Concur_CountryMap", "Country Map (left block)", "adp_country (3-char)", "concur_country; currency_code", "Legal Country (both); Reimbursement Currency (UKG)"),
 ("Org", "ADP_Concur_OrgMap", "Org Map", "business_unit_desc; home_department_code", "org_unit_1, default_language, currency (by BU); org_unit_2 (by dept code)", "ADP Org Unit 1, Org Unit 2, Locale fallback, Reimbursement Currency"),
 ("Language", "ADP_Concur_LanguageMap", "Language Map (A:C)", "language_desc", "adp_language (stem 'xx_')", "ADP Locale Code"),
 ("Locale", "ADP_Concur_LocaleMap", "Language Map (M:O)", "country_code (2-char)", "locale_code", "UKG Locale Code"),
 ("Salary", "ADP_Concur_SalaryMap", "Salary Map", "pay_grade_code (trimmed)", "expense_map, travel_map, invoice_map", "Concur Profile, Travel Profile, Invoice Limit, and ADP Invoice Access"),
 ("Supervisor", "ADP_Concur_SupervisorMap", "Supervisor Map", "file_number (the employee)", "supervisor_id", "Supervisor ID override (wins over the export)"),
 ("Invoice", "ADP_Concur_InvoiceMap", "Invoice Exception Map", "file_number (trimmed)", "access (Y/N)", "Invoice Access override"),
 ("Company", "ADP_Concur_CompanyMap", "UKG Company Map", "site_location_code", "concur_company_code, expense_group_code, ledger_code, custom_5_code", "UKG Org Unit 1, Custom 21, Ledger Code, Custom 5"),
 ("CountryRef", "ADP_Concur_CountryRef", "Country Map (right block)", "country_code (2-char)", "currency_code", "Reference only"),
 ("Role", "ADP_Concur_RoleMap", "Role Map", "role", "automatic (Y/N)", "Reference only; the record builders do not read it"),
]
rows = []
for short, tbl, tab, key, ret, used in maps:
    rows.append([f"Map - {short}", tbl, tab, key, ret, used, f"=COUNTA('Map - {short}'!A:A)-3"])
sheet("Map Tables", "Map Tables",
      "Each lookup table: the workbook tab it is loaded from, the key it is looked up by, and what it returns. Rows are counted from the snapshot tabs.",
      ["Guide Tab", "Database Table", "Workbook Tab", "Lookup Key", "Returns", "Used For", "Rows"],
      rows, [18, 28, 26, 32, 46, 48, 8])

skip = {"map_key", "ref_key", "locale_key", "row_state"}
for short, tbl, tab, *_ in maps:
    cur = conn.execute(f"SELECT * FROM {tbl} ORDER BY 1")
    cols = [d[0] for d in cur.description if d[0] not in skip]
    data = [[r[c] for c in cols] for r in cur]
    sheet(f"Map - {short}", f"{tab}  ->  {tbl}", f"Snapshot of the live table taken {date.today():%Y-%m-%d}. Edit maps in the app, not here.",
          cols, data, [max(12, min(40, max([len(c)] + [len(str(x or "")) for x in (row[i] for row in data)]) + 2))
                       for i, c in enumerate(cols)])

# ------------------------------------------------------------ 7 config
D0 = M.DEFAULT_CONFIG
def j(v): return json.dumps(v) if not isinstance(v, str) or v == "" else v
cfg_rows = []
explain = {
 "login_id": "How the Login ID is built. See Derived Fields: login_id.",
 "login_id_by_source": "Per-source overrides merged over login_id.",
 "password": "The 305 password (G). Ignored while Password Generation = SSO.",
 "records_by_source": "Which record types each source can write.",
 "scope": "Who each record type covers: all | invoice (Invoice Access = Y) | off.",
 "blank_fields": "Record positions forced empty whatever the mapping says, e.g. {\"305\": [\"Z\"]}.",
 "extract": "File shape: delimiter, quoting, line ending, order, file name, output folder.",
 "block_on": "Which problems keep an employee out of the file (true) or only warn (false).",
 "import_settings": "The 100 record. See Record 100.",
}
for k, v in cfg.items():
    if isinstance(v, dict) and k not in ("login_id_by_source",) and v:
        for sk, sv in v.items():
            dv = D0.get(k, {}).get(sk) if isinstance(D0.get(k), dict) else None
            cfg_rows.append([k, sk, j(sv), j(dv), explain.get(k, "")])
    else:
        cfg_rows.append([k, "", j(v), j(D0.get(k)), explain.get(k, "")])
sheet("Config", "Config - ADP_Concur_Config.json",
      "Current values against the code defaults in ADP_Concur_Map.DEFAULT_CONFIG. Yellow = the file overrides the default.",
      ["Setting", "Key", "Current Value", "Code Default", "What It Does"], cfg_rows, [22, 26, 40, 40, 70],
      [DIFF if r_[2] != r_[3] else None for r_ in cfg_rows])

# ------------------------------------------------------------ 8 validation
b = cfg["block_on"]
def sev(flag): return "Error (excluded)" if b.get(flag) else "Warning"
val = [
 ("file_number", "No File Number", sev("missing_employee_id"), "block_on.missing_employee_id"),
 ("login_id", "No Login ID (no source value for the rule)", sev("missing_login_id"), "block_on.missing_login_id"),
 ("work_email", "No work email in the export", "Warning", "-"),
 ("<any derived>", "A value contains a '... Not Mapped' / '... Not Matched' text. Terminated people are included too, because Concur checks the hierarchy node before status.", sev("unmapped"), "block_on.unmapped"),
 ("supervisor_id", "An active employee with no supervisor", "Warning", "-"),
 ("term_date", "Inactive (Status N) but no Termination Date", "Warning", "-"),
 ("duplicate_rows", "More than one export row for the File Number, or the person is in both ADP and UKG", "Warning", "-"),
 ("invoice_limit", "Invoice Access Y (from the exception map) but the grade's limit is 0", "Warning", "-"),
 ("org_unit_2", "Org Unit 2 is a name, not a numeric code. Concur's connected list rejects it.", sev("org_unit_2_code"), "block_on.org_unit_2_code"),
 ("supervisor_chain", "The supervisor is not in the load, the person reports to themselves, or the chain loops", sev("broken_supervisor"), "block_on.broken_supervisor"),
 ("supervisor_chain", "The supervisor comes from the other system (ADP <-> UKG)", "Warning", "-"),
 ("supervisor_chain", "An active employee's approver is inactive", "Warning", "-"),
 ("supervisor_chain", "The supervisor is deleted or held out of the 305", "Warning", "-"),
 ("password", "Passwords run in an unbroken fill series (5 or more, e.g. Welcome01..Welcome82)", "Warning", "-"),
]
sheet("Validation", "Validation - Exceptions",
      "Rebuilt on every derive into ADP_Concur_Exceptions. Any Error keeps that employee out of every record type in the extract.",
      ["Field", "Condition", "Current Severity", "Controlled By"], [list(v) for v in val], [20, 90, 18, 28],
      [DIFF if v[2].startswith("Error") else None for v in val])

order = ["Overview", "Source Columns", "Derived Fields", "Record 100", "Record 305", "Record 350", "Record 360",
         "Record 700", "Map Tables", "Config", "Validation"]
wb._sheets = [wb[n] for n in order] + [s_ for s_ in wb._sheets if s_.title not in order]
wb.save(OUT)
print("saved", OUT)
