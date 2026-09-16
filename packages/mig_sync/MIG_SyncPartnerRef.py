# MIG_SyncPartnerRef.py
"""
MIG_SyncPartnerRef.py
---------------------
Syncs partner reference data (CRS945) from a SOURCE tenant to a DESTINATION
tenant via CRS945MI.UpdPartnerRef and CRS945MI.AddPartnerRef.

Steps
-----
1.  SOURCE tenant -> CRS945MI.LstPartnerRef (SOURCE DIVI) -> records in memory.
2.  DEST   tenant -> CRS945MI.LstPartnerRef (DEST   DIVI) -> records in memory.
3.  Diff: find records where PRF1/PRF2/TX15/TX40 differ (matched on
    DONR+DOVA+MEPF), and records missing from DEST entirely.
4.  On DEST:
      a. CRS945MI.UpdPartnerRef - update records that exist but differ.
      b. CRS945MI.AddPartnerRef - add records that are missing from DEST.

Ported to a library used by MIG_App.py: input()/print() are gone, tqdm is
replaced with a job progress(done, total, message) callback, and tenant auth
is an explicit MIG_Api.Tenant.
"""

from __future__ import annotations

import datetime
from collections import Counter
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

import MIG_Api

# =============================================================================
# Constants
# =============================================================================

PARTNER_REF_COLS = ["DIVI", "DONR", "DOVA", "MEPF", "PRF1", "PRF2", "TX15", "TX40"]
KEY_FIELDS = ["DIVI", "DONR", "DOVA", "MEPF", "PRF1", "PRF2"]
VALUE_FIELDS = ["TX15", "TX40"]

DEFAULT_BATCH_SIZE = 100


def _noop_progress(done, total, message=""):
    pass


# =============================================================================
# CRS945MI.LstPartnerRef -> list of dicts
# =============================================================================

def fetch_partner_refs(tenant: MIG_Api.Tenant, session: requests.Session,
                        divi: str, label: str) -> list[dict]:
    """CRS945MI.LstPartnerRef for the given division."""
    payload = {
        "program": "CRS945MI",
        "transactions": [{
            "transaction": "LstPartnerRef",
            "record": {"DIVI": divi},
            "selectedColumns": PARTNER_REF_COLS,
        }],
    }
    data = MIG_Api.post_to_m3(tenant, payload, session)

    rows: list[dict] = []
    for result in data.get("results", []):
        for record in result.get("records", []):
            rows.append({col: (record.get(col, "") or "").strip() for col in PARTNER_REF_COLS})
    return rows


# =============================================================================
# Diff logic
# =============================================================================

def _row_key(row: dict) -> tuple:
    return tuple(row.get(f, "") for f in KEY_FIELDS)


def diff_partner_refs(source_rows: list[dict],
                       dest_rows: list[dict]) -> tuple[list[dict], list[dict]]:
    """
    Compare SOURCE and DEST partner-ref lists.

    Returns (to_update, to_add):
        to_update - rows whose key exists in DEST but whose value fields differ
        to_add    - rows whose key is entirely absent from DEST
    """
    dest_index: dict[tuple, dict] = {_row_key(r): r for r in dest_rows}

    to_update: list[dict] = []
    to_add: list[dict] = []
    for src in source_rows:
        key = _row_key(src)
        dest = dest_index.get(key)
        if dest is None:
            to_add.append(src)
        elif any(src.get(f, "") != dest.get(f, "") for f in VALUE_FIELDS):
            to_update.append(src)

    return to_update, to_add


def preview_partner_refs(to_update: list[dict], to_add: list[dict],
                          dest_rows: list[dict]) -> list[dict]:
    """Annotate the diff plan for display: ACTION (UPDATE/ADD) + CHANGES."""
    dest_index: dict[tuple, dict] = {_row_key(r): r for r in dest_rows}

    def _changes_summary(src_row: dict) -> str:
        dest_row = dest_index.get(_row_key(src_row), {})
        parts = []
        for f in VALUE_FIELDS:
            sv, dv = src_row.get(f, ""), dest_row.get(f, "")
            if sv != dv:
                parts.append(f'{f}: "{dv}" -> "{sv}"')
        return " | ".join(parts)

    return ([{**row, "ACTION": "UPDATE", "CHANGES": _changes_summary(row)} for row in to_update]
            + [{**row, "ACTION": "ADD", "CHANGES": ""} for row in to_add])


# =============================================================================
# Build API records
# =============================================================================

def build_upd_record(src_row: dict, dest_divi: str) -> dict:
    """Record for CRS945MI.UpdPartnerRef using SOURCE values + DEST DIVI."""
    rec = {col: src_row.get(col, "") for col in PARTNER_REF_COLS}
    rec["DIVI"] = dest_divi
    return {k: v for k, v in rec.items() if v != ""}


def build_add_record(src_row: dict, dest_divi: str) -> dict:
    """Record for CRS945MI.AddPartnerRef using SOURCE values + DEST DIVI."""
    rec = {col: src_row.get(col, "") for col in PARTNER_REF_COLS}
    rec["DIVI"] = dest_divi
    return {k: v for k, v in rec.items() if v != ""}


# =============================================================================
# CRS945MI.UpdPartnerRef / AddPartnerRef on DEST (batched)
# =============================================================================

def _run_batched(tenant: MIG_Api.Tenant, transaction: str, records: list[dict],
                  session: requests.Session, batch_size: int,
                  progress) -> tuple[list[dict], list[tuple[dict, str]]]:
    successes: list[dict] = []
    failures: list[tuple[dict, str]] = []
    total = len(records)
    label = f"CRS945MI.{transaction}"
    progress(0, total, label)

    for i in range(0, total, batch_size):
        batch = records[i:i + batch_size]
        payload = {
            "program": "CRS945MI",
            "transactions": [{"transaction": transaction, "record": rec,
                               "selectedColumns": PARTNER_REF_COLS} for rec in batch],
        }
        try:
            result = MIG_Api.post_to_m3(tenant, payload, session)
            api_results = result.get("results", [])
            for j, res in enumerate(api_results):
                err = (res.get("errorMessage") or "").strip() if isinstance(res, dict) else ""
                rec = batch[j] if j < len(batch) else {}
                if err:
                    failures.append((rec, err))
                else:
                    successes.append(rec)
            for k in range(len(api_results), len(batch)):
                failures.append((batch[k], "No result returned"))
        except Exception as exc:
            for rec in batch:
                failures.append((rec, str(exc)))
        progress(min(i + batch_size, total), total, label)

    return successes, failures


def run_upd_batched(tenant: MIG_Api.Tenant, records: list[dict],
                     session: requests.Session, batch_size: int,
                     progress=_noop_progress) -> tuple[list[dict], list[tuple[dict, str]]]:
    return _run_batched(tenant, "UpdPartnerRef", records, session, batch_size, progress)


def run_add_batched(tenant: MIG_Api.Tenant, records: list[dict],
                     session: requests.Session, batch_size: int,
                     progress=_noop_progress) -> tuple[list[dict], list[tuple[dict, str]]]:
    return _run_batched(tenant, "AddPartnerRef", records, session, batch_size, progress)


# =============================================================================
# Reporting
# =============================================================================

def summary_counts(successes: list[dict], failures: list[tuple[dict, str]]) -> dict:
    """{'OK': n, '<error message>': n, ...} - same counting the CLI printed."""
    counts: Counter = Counter()
    counts["OK"] = len(successes)
    for _, err in failures:
        counts[err] += 1
    return dict(counts.most_common())


def export_errors_xlsx(
    upd_successes: list[dict], upd_failures: list[tuple[dict, str]],
    add_successes: list[dict], add_failures: list[tuple[dict, str]],
    out_dir: Path,
) -> Path:
    """
    Sheets: Summary (error-message totals for both steps), UpdPartnerRef Err,
    AddPartnerRef Err (one row per failed record). Returns the written path.
    """
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"MIG_SyncPartnerRef_Errors_{ts}.xlsx"

    wb = Workbook()
    hdr_font = Font(name="Arial", bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", start_color="2F5496")
    ok_fill = PatternFill("solid", start_color="E2EFDA")
    err_fill = PatternFill("solid", start_color="FCE4D6")
    ctr = Alignment(horizontal="center")

    def _write_error_sheet(ws, fields, failures):
        for col, f in enumerate(fields, start=1):
            c = ws.cell(row=1, column=col, value=f)
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = ctr
        for row_i, (rec, err) in enumerate(failures, start=2):
            for col, f in enumerate(fields, start=1):
                val = err if f == "ERROR" else rec.get(f, "")
                c = ws.cell(row=row_i, column=col, value=val)
                c.font = Font(name="Arial")
        for col, f in enumerate(fields, start=1):
            max_len = max(len(f), *(len(str(rec.get(f, "") if f != "ERROR" else err))
                                     for rec, err in failures))
            ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = min(max_len + 2, 60)

    ws_sum = wb.active
    ws_sum.title = "Summary"
    for col, heading in enumerate(["Step", "Message", "Count"], start=1):
        c = ws_sum.cell(row=1, column=col, value=heading)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = ctr

    row_i = 2
    for step_label, successes, failures in [
        ("UpdPartnerRef", upd_successes, upd_failures),
        ("AddPartnerRef", add_successes, add_failures),
    ]:
        for msg, count in summary_counts(successes, failures).items():
            fill = ok_fill if msg == "OK" else err_fill
            for col, val in enumerate([step_label, msg, count], start=1):
                c = ws_sum.cell(row=row_i, column=col, value=val)
                c.fill = fill
                c.font = Font(name="Arial")
                if col == 3:
                    c.alignment = ctr
            row_i += 1

    ws_sum.column_dimensions["A"].width = 20
    ws_sum.column_dimensions["B"].width = 80
    ws_sum.column_dimensions["C"].width = 10

    if upd_failures:
        _write_error_sheet(wb.create_sheet("UpdPartnerRef Err"), PARTNER_REF_COLS + ["ERROR"], upd_failures)
    if add_failures:
        _write_error_sheet(wb.create_sheet("AddPartnerRef Err"), PARTNER_REF_COLS + ["ERROR"], add_failures)

    wb.save(out_path)
    return out_path
