"""
MIG_App - Flask front end for the MIG Sync package.

Nine tabs, one per original MIG_Sync*.py / MIG_*.py CLI script: Panel Views,
Partner Media, Partner Ref, Sorting Options, Sorting Orders, Transl Data,
Field Groups, Export Summary, Migration Review.

This is a THIN WEB WRAPPER, not a persistent-database app like m3_security.
There is no SQLite capture of the sync data itself: every run does a live
SOURCE -> DEST comparison in memory, exactly like the CLI scripts did, with
input() replaced by a browser UI and print() replaced by JSON + a job-progress
poll (see start_job() below). Tenants (SOURCE / DEST / single) are held
server-side in an in-memory dict keyed by a uuid4 tenant_id - nothing about a
run, or a connected tenant, survives a restart of this process.

See MIG_README.md for the full per-tab reference.
"""

from __future__ import annotations

import argparse
import logging
import threading
import time
import traceback
import uuid
from pathlib import Path

import requests
from flask import Flask, jsonify, render_template, request, send_file

import MIG_Api
import MIG_ExportSummary
import MIG_GenerateFieldGroups
import MIG_MigrationReview
import MIG_SyncPanelViews
import MIG_SyncPartnerMedia
import MIG_SyncPartnerRef
import MIG_SyncSortingOptions
import MIG_SyncSortingOrders
import MIG_SyncTranslData

log = logging.getLogger("MIG_App")

BASE_DIR = Path(__file__).parent.resolve()
UPLOAD_DIR = BASE_DIR / "input" / "mig_sync"
OUTPUT_DIR = BASE_DIR / "output" / "mig_sync"

# ---------------------------------------------------------------------------
# Shared repo-root resources. These are NOT package-local - do not "fix" them
# to point inside packages/mig_sync/. Every Infor automation tool in this repo
# reads .ionapi files from the same ionapi/ folder, and an external process
# outside this repo watches evs100/ToProcess/ and moves what it consumes into
# evs100/Complete/. Pointing either of these at a package-local copy breaks
# the real M3 import pickup.
# ---------------------------------------------------------------------------
IONAPI_DIR = BASE_DIR.parent.parent / "ionapi"
EVS100_TO_PROCESS = BASE_DIR.parent.parent / "evs100" / "ToProcess"

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config["MAX_CONTENT_LENGTH"] = 256 * 1024 * 1024


@app.errorhandler(Exception)
def _handle(exc):
    # Flask's own HTTP errors carry a numeric .code; pass those straight
    # through rather than dressing them up as a 500 with a traceback.
    if hasattr(exc, "code") and isinstance(getattr(exc, "code"), int):
        return jsonify(status="error", message=str(exc)), exc.code
    log.error("%s", traceback.format_exc())
    return jsonify(status="error", message=str(exc),
                   trace=traceback.format_exc()), 500


@app.errorhandler(MIG_Api.M3ApiError)
def _handle_m3(exc):
    return jsonify(status="error", message=str(exc)), 400


@app.route("/")
def index():
    return render_template("MIG_Index.html")


# ------------------------------------------------------------ job plumbing
#
# Long M3 runs (or anything that loops over many records) happen on a worker
# thread; the browser polls /api/job/<id>. Copied from the same pattern in
# M3_Security_App.py / ADP_Concur_App.py.

_JOBS: dict[str, dict] = {}
_JOBS_LOCK = threading.Lock()
_JOB_KEEP = 40


def _job_set(job_id: str, **fields) -> None:
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        if job is not None:
            job.update(fields)


def start_job(label: str, fn) -> str:
    """
    Run fn(progress) on a worker thread. progress(done, total, message)
    updates what /api/job/<id> reports. fn's return value becomes the job's
    "result" once it finishes.
    """
    job_id = str(uuid.uuid4())
    with _JOBS_LOCK:
        for old in sorted(_JOBS, key=lambda k: _JOBS[k]["started_at"])[:-_JOB_KEEP]:
            if _JOBS[old]["status"] != "running":
                _JOBS.pop(old, None)
        _JOBS[job_id] = {"id": job_id, "label": label, "status": "running",
                          "done": 0, "total": 0, "phase": "", "result": None,
                          "error": None, "started_at": time.time(),
                          "finished_at": None}

    def progress(done: int, total: int, message: str = "") -> None:
        _job_set(job_id, done=done, total=total, phase=message or "")

    def run() -> None:
        try:
            result = fn(progress)
            _job_set(job_id, status="done", result=result, finished_at=time.time())
        except Exception as exc:
            log.error("job %s failed: %s", label, traceback.format_exc())
            _job_set(job_id, status="error", error=str(exc), finished_at=time.time())

    threading.Thread(target=run, name=f"mig-{label}", daemon=True).start()
    return job_id


@app.route("/api/job/<job_id>")
def api_job(job_id: str):
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        snapshot = dict(job) if job else None
    if snapshot is None:
        return jsonify(status="error", message="Unknown job."), 404
    snapshot.pop("started_at", None)
    snapshot.pop("finished_at", None)
    return jsonify(status="success", job=snapshot)


def _job_result(job_id: str) -> dict:
    """The stored result of a finished job, or raise M3ApiError."""
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if job is None:
        raise MIG_Api.M3ApiError("Unknown job.")
    if job["status"] != "done":
        raise MIG_Api.M3ApiError(f"Job is not finished yet (status={job['status']}).")
    return job["result"]


# --------------------------------------------------------- tenant plumbing
#
# One /api/tenant/connect route, shared by every tab that needs a SOURCE
# and/or DEST tenant. Tenants live server-side in this dict, keyed by a
# fresh uuid4 - the access_token never goes to the browser.

_TENANTS: dict[str, MIG_Api.Tenant] = {}
_TENANTS_LOCK = threading.Lock()


def _tenant(tenant_id: str) -> MIG_Api.Tenant:
    with _TENANTS_LOCK:
        tenant = _TENANTS.get((tenant_id or "").strip())
    if tenant is None:
        raise MIG_Api.M3ApiError("Unknown or expired tenant_id - connect again.")
    return tenant


@app.route("/api/ionapi")
def api_ionapi():
    return jsonify(status="success", files=MIG_Api.list_ionapi_files(IONAPI_DIR))


@app.route("/api/tenant/connect", methods=["POST"])
def api_tenant_connect():
    data = request.get_json(force=True) or {}
    ionapi_file = (data.get("ionapi_file") or "").strip()
    if not ionapi_file:
        return jsonify(status="error", message="An .ionapi file is required."), 400

    # Sanitize to a bare filename inside IONAPI_DIR - never let the browser
    # point this at an arbitrary path.
    ionapi_path = IONAPI_DIR / Path(ionapi_file).name
    if not ionapi_path.is_file():
        return jsonify(status="error", message=f"{ionapi_path.name} not found in {IONAPI_DIR}."), 400

    tenant = MIG_Api.Tenant(
        ionapi_path=ionapi_path,
        label=(data.get("label") or "").strip(),
        company=(data.get("company") or "").strip(),
        division=(data.get("division") or "").strip(),
        global_scope=bool(data.get("global_scope")),
    )
    with requests.Session() as session:
        MIG_Api.authenticate(tenant, session)

    tenant_id = str(uuid.uuid4())
    with _TENANTS_LOCK:
        _TENANTS[tenant_id] = tenant

    return jsonify(status="success", tenant_id=tenant_id, ti=tenant.ti, iu=tenant.iu,
                   company=tenant.company, division=tenant.division,
                   ionapi_file=ionapi_path.name,
                   message=f"Connected: {ionapi_path.name} (CONO={tenant.company or '-'} "
                           f"DIVI={tenant.division or '-'}).")


# ------------------------------------------------------------- downloads
#
# One safe-download route shared by every tab, mirroring the /api/export/<kind>
# pattern in M3_Security_App.py. `kind` selects the directory; the filename is
# always reduced to its bare name first, so this can never escape the folder.

_DOWNLOAD_DIRS = {
    "evs100": EVS100_TO_PROCESS,
    "partner_media": OUTPUT_DIR / "partner_media",
    "partner_ref": OUTPUT_DIR / "partner_ref",
    "transl_data": OUTPUT_DIR / "transl_data",
    "export_summary": MIG_ExportSummary.DEFAULT_OUTPUT_FOLDER,
    "migration_review": MIG_MigrationReview.OUTPUT_DIR,
}


@app.route("/api/download/<kind>/<path:filename>")
def api_download(kind: str, filename: str):
    directory = _DOWNLOAD_DIRS.get(kind)
    if directory is None:
        return jsonify(status="error", message=f"Unknown download kind '{kind}'."), 400
    path = directory / Path(filename).name
    if not path.is_file():
        return jsonify(status="error", message=f"{path.name} not found."), 404
    return send_file(str(path), as_attachment=True, download_name=path.name)


# =============================================================================
# Panel Views  (SOURCE + DEST, optional PGNM filter, EVS100 file)
# =============================================================================

@app.route("/api/panel-views/fetch-diff", methods=["POST"])
def api_panel_views_fetch_diff():
    data = request.get_json(force=True) or {}
    source = _tenant(data.get("source_tenant_id"))
    dest = _tenant(data.get("dest_tenant_id"))
    pgnm_filter = (data.get("pgnm_filter") or "").strip().upper()
    query = MIG_SyncPanelViews._build_exportmi_query(pgnm_filter)

    def work(progress):
        progress(0, 2, "EXPORTMI.Select CSYSPV (SOURCE)")
        with requests.Session() as session:
            source_views = MIG_SyncPanelViews.fetch_panel_views(source, session, "SOURCE", query)
        progress(1, 2, "EXPORTMI.Select CSYSPV (DEST)")
        with requests.Session() as session:
            dest_views = MIG_SyncPanelViews.fetch_panel_views(dest, session, "DEST", query)
        progress(2, 2, "Comparing SOURCE vs DEST")
        result = MIG_SyncPanelViews.diff_panel_views(source_views, dest_views)
        result["pgnm_filter"] = pgnm_filter
        return result

    return jsonify(status="success", job_id=start_job("Panel Views fetch+diff", work))


@app.route("/api/panel-views/export", methods=["POST"])
def api_panel_views_export():
    data = request.get_json(force=True) or {}
    result = _job_result(data.get("job_id"))
    path = MIG_SyncPanelViews.export_evs100_xlsx(result["differing"], EVS100_TO_PROCESS)
    return jsonify(status="success", file=path.name, download="evs100/" + path.name,
                   message=f"{len(result['differing'])} view(s) exported to {path.name}.")


@app.route("/api/panel-views/upload", methods=["POST"])
def api_panel_views_upload():
    data = request.get_json(force=True) or {}
    dest = _tenant(data.get("dest_tenant_id"))
    filename = Path((data.get("filename") or "")).name
    path = EVS100_TO_PROCESS / filename
    if not path.is_file():
        return jsonify(status="error", message=f"{filename} not found."), 404

    def work(progress):
        progress(0, 1, f"Uploading {filename}")
        with requests.Session() as session:
            ok, message = MIG_Api.upload_file_to_m3(dest, path, session)
        progress(1, 1, message)
        if not ok:
            raise MIG_Api.M3ApiError(message)
        return {"ok": ok, "message": message, "file": filename}

    return jsonify(status="success", job_id=start_job("Panel Views upload", work))


@app.route("/api/panel-views/process-in-m3", methods=["POST"])
def api_panel_views_process():
    data = request.get_json(force=True) or {}
    dest = _tenant(data.get("dest_tenant_id"))
    filename = Path((data.get("filename") or "")).name

    def work(progress):
        progress(0, 1, f"EVS100MI.ImportFile {filename}")
        with requests.Session() as session:
            ok, message = MIG_Api.process_file_in_m3(dest, filename, session)
        progress(1, 1, message)
        if not ok:
            raise MIG_Api.M3ApiError(message)
        return {"ok": ok, "message": message, "file": filename}

    return jsonify(status="success", job_id=start_job("Panel Views process", work))


# =============================================================================
# Partner Media  (SOURCE + DEST, direct write, always-generated summary xlsx)
# =============================================================================

@app.route("/api/partner-media/fetch-diff", methods=["POST"])
def api_partner_media_fetch_diff():
    data = request.get_json(force=True) or {}
    source = _tenant(data.get("source_tenant_id"))
    dest = _tenant(data.get("dest_tenant_id"))

    def work(progress):
        progress(0, 2, "CRS949MI.LstPartnerMedia (SOURCE)")
        with requests.Session() as session:
            source_rows = MIG_SyncPartnerMedia.fetch_partner_media(
                source, session, source.division, "SOURCE")
        progress(1, 2, "CRS949MI.LstPartnerMedia (DEST)")
        with requests.Session() as session:
            dest_rows = MIG_SyncPartnerMedia.fetch_partner_media(
                dest, session, dest.division, "DEST")
        progress(2, 2, "Comparing SOURCE vs DEST")
        to_update, to_add = MIG_SyncPartnerMedia.diff_partner_media(source_rows, dest_rows)
        preview = MIG_SyncPartnerMedia.preview_partner_media(to_update, to_add, dest_rows)
        return {
            "to_update": to_update, "to_add": to_add, "dest_rows": dest_rows,
            "preview": preview,
            "counts": {"source": len(source_rows), "dest": len(dest_rows),
                       "to_update": len(to_update), "to_add": len(to_add)},
        }

    return jsonify(status="success", job_id=start_job("Partner Media fetch+diff", work))


@app.route("/api/partner-media/apply", methods=["POST"])
def api_partner_media_apply():
    data = request.get_json(force=True) or {}
    dest = _tenant(data.get("dest_tenant_id"))
    prior = _job_result(data.get("job_id"))
    to_update, to_add, dest_rows = prior["to_update"], prior["to_add"], prior["dest_rows"]
    dest_divi = dest.division
    batch_size = MIG_SyncPartnerMedia.BATCH_SIZE

    def work(progress):
        dest_index = {MIG_SyncPartnerMedia._row_key(r): r for r in dest_rows}
        to_update_ok = [r for r in to_update if MIG_SyncPartnerMedia.resolve_add_transaction(r)]
        to_add_ok = [r for r in to_add if MIG_SyncPartnerMedia.resolve_add_transaction(r)]

        dlt_records = [MIG_SyncPartnerMedia.build_dlt_record(dest_index[MIG_SyncPartnerMedia._row_key(r)], dest_divi)
                        for r in to_update_ok]
        upd_add_records = [MIG_SyncPartnerMedia.build_add_record(r, dest_divi) for r in to_update_ok]
        add_records = [MIG_SyncPartnerMedia.build_add_record(r, dest_divi) for r in to_add]

        dlt_successes, dlt_failures = [], []
        upd_add_successes, upd_add_failures, upd_add_skipped = [], [], []
        add_successes, add_failures, add_skipped = [], [], []

        with requests.Session() as session:
            if dlt_records:
                dlt_successes, dlt_failures = MIG_SyncPartnerMedia.run_dlt_batched(
                    dest, dlt_records, session, batch_size, progress=progress)

            if dlt_successes and upd_add_records:
                dlt_success_keys = {MIG_SyncPartnerMedia._row_key(r) for r in dlt_successes}
                records_to_readd = [
                    rec for rec, src in zip(upd_add_records, to_update_ok)
                    if MIG_SyncPartnerMedia._row_key(
                        MIG_SyncPartnerMedia.build_dlt_record(dest_index[MIG_SyncPartnerMedia._row_key(src)], dest_divi)
                    ) in dlt_success_keys
                ]
                if records_to_readd:
                    upd_add_successes, upd_add_failures, upd_add_skipped = MIG_SyncPartnerMedia.run_add_batched(
                        dest, records_to_readd, session, batch_size, progress=progress, desc="update")

            if add_records:
                add_successes, add_failures, add_skipped = MIG_SyncPartnerMedia.run_add_batched(
                    dest, add_records, session, batch_size, progress=progress, desc="new")

        out_dir = OUTPUT_DIR / "partner_media"
        xlsx_path = MIG_SyncPartnerMedia.export_summary_xlsx(
            prior["counts"]["source"], prior["counts"]["dest"],
            dlt_successes, dlt_failures,
            upd_add_successes, upd_add_failures, upd_add_skipped,
            add_successes, add_failures, add_skipped,
            out_dir)

        return {
            "summary_file": xlsx_path.name, "download": "partner_media/" + xlsx_path.name,
            "counts": {
                "dlt_success": len(dlt_successes), "dlt_failed": len(dlt_failures),
                "update_success": len(upd_add_successes), "update_failed": len(upd_add_failures),
                "update_skipped": len(upd_add_skipped),
                "add_success": len(add_successes), "add_failed": len(add_failures),
                "add_skipped": len(add_skipped),
            },
        }

    return jsonify(status="success", job_id=start_job("Partner Media apply", work))


# =============================================================================
# Partner Ref  (SOURCE + DEST, direct write, error xlsx only on failures)
# =============================================================================

@app.route("/api/partner-ref/fetch-diff", methods=["POST"])
def api_partner_ref_fetch_diff():
    data = request.get_json(force=True) or {}
    source = _tenant(data.get("source_tenant_id"))
    dest = _tenant(data.get("dest_tenant_id"))

    def work(progress):
        progress(0, 2, "CRS945MI.LstPartnerRef (SOURCE)")
        with requests.Session() as session:
            source_rows = MIG_SyncPartnerRef.fetch_partner_refs(source, session, source.division, "SOURCE")
        progress(1, 2, "CRS945MI.LstPartnerRef (DEST)")
        with requests.Session() as session:
            dest_rows = MIG_SyncPartnerRef.fetch_partner_refs(dest, session, dest.division, "DEST")
        progress(2, 2, "Comparing SOURCE vs DEST")
        to_update, to_add = MIG_SyncPartnerRef.diff_partner_refs(source_rows, dest_rows)
        preview = MIG_SyncPartnerRef.preview_partner_refs(to_update, to_add, dest_rows)
        return {
            "to_update": to_update, "to_add": to_add, "preview": preview,
            "counts": {"source": len(source_rows), "dest": len(dest_rows),
                       "to_update": len(to_update), "to_add": len(to_add)},
        }

    return jsonify(status="success", job_id=start_job("Partner Ref fetch+diff", work))


@app.route("/api/partner-ref/apply", methods=["POST"])
def api_partner_ref_apply():
    data = request.get_json(force=True) or {}
    dest = _tenant(data.get("dest_tenant_id"))
    prior = _job_result(data.get("job_id"))
    dest_divi = dest.division
    batch_size = MIG_SyncPartnerRef.DEFAULT_BATCH_SIZE

    upd_records = [MIG_SyncPartnerRef.build_upd_record(r, dest_divi) for r in prior["to_update"]]
    add_records = [MIG_SyncPartnerRef.build_add_record(r, dest_divi) for r in prior["to_add"]]

    def work(progress):
        upd_successes, upd_failures = [], []
        add_successes, add_failures = [], []
        with requests.Session() as session:
            if upd_records:
                upd_successes, upd_failures = MIG_SyncPartnerRef.run_upd_batched(
                    dest, upd_records, session, batch_size, progress=progress)
            if add_records:
                add_successes, add_failures = MIG_SyncPartnerRef.run_add_batched(
                    dest, add_records, session, batch_size, progress=progress)

        result = {
            "counts": {"update_success": len(upd_successes), "update_failed": len(upd_failures),
                       "add_success": len(add_successes), "add_failed": len(add_failures)},
            "summary": {
                "UpdPartnerRef": MIG_SyncPartnerRef.summary_counts(upd_successes, upd_failures),
                "AddPartnerRef": MIG_SyncPartnerRef.summary_counts(add_successes, add_failures),
            },
        }
        if upd_failures or add_failures:
            xlsx_path = MIG_SyncPartnerRef.export_errors_xlsx(
                upd_successes, upd_failures, add_successes, add_failures,
                OUTPUT_DIR / "partner_ref")
            result["error_file"] = xlsx_path.name
            result["download"] = "partner_ref/" + xlsx_path.name
        return result

    return jsonify(status="success", job_id=start_job("Partner Ref apply", work))


# =============================================================================
# Sorting Options  (SOURCE + DEST, EVS100 file)
# =============================================================================

@app.route("/api/sorting-options/fetch-diff", methods=["POST"])
def api_sorting_options_fetch_diff():
    data = request.get_json(force=True) or {}
    source = _tenant(data.get("source_tenant_id"))
    dest = _tenant(data.get("dest_tenant_id"))

    def work(progress):
        progress(0, 2, "CRS021MI.LstSrtOpt (SOURCE)")
        with requests.Session() as session:
            source_all = MIG_SyncSortingOptions.list_all_srt_opts(source, session, "SOURCE")
        progress(1, 2, "CRS021MI.LstSrtOpt (DEST)")
        with requests.Session() as session:
            dest_all = MIG_SyncSortingOptions.list_all_srt_opts(dest, session, "DEST")
        progress(2, 2, "Comparing SOURCE vs DEST")
        return MIG_SyncSortingOptions.diff_sorting_options(source_all, dest_all)

    return jsonify(status="success", job_id=start_job("Sorting Options fetch+diff", work))


@app.route("/api/sorting-options/export", methods=["POST"])
def api_sorting_options_export():
    data = request.get_json(force=True) or {}
    result = _job_result(data.get("job_id"))
    add_records, act_records, std_records = MIG_SyncSortingOptions.build_record_sets(result["missing"])
    path = MIG_SyncSortingOptions.export_evs100_xlsx(add_records, act_records, std_records, EVS100_TO_PROCESS)
    return jsonify(status="success", file=path.name, download="evs100/" + path.name,
                   counts={"add": len(add_records), "activate": len(act_records), "standard": len(std_records)},
                   message=f"{len(result['missing'])} record(s) exported to {path.name}.")


@app.route("/api/sorting-options/upload", methods=["POST"])
def api_sorting_options_upload():
    data = request.get_json(force=True) or {}
    dest = _tenant(data.get("dest_tenant_id"))
    filename = Path((data.get("filename") or "")).name
    path = EVS100_TO_PROCESS / filename
    if not path.is_file():
        return jsonify(status="error", message=f"{filename} not found."), 404

    def work(progress):
        progress(0, 1, f"Uploading {filename}")
        with requests.Session() as session:
            ok, message = MIG_Api.upload_file_to_m3(dest, path, session)
        progress(1, 1, message)
        if not ok:
            raise MIG_Api.M3ApiError(message)
        return {"ok": ok, "message": message, "file": filename}

    return jsonify(status="success", job_id=start_job("Sorting Options upload", work))


@app.route("/api/sorting-options/process-in-m3", methods=["POST"])
def api_sorting_options_process():
    data = request.get_json(force=True) or {}
    dest = _tenant(data.get("dest_tenant_id"))
    filename = Path((data.get("filename") or "")).name

    def work(progress):
        progress(0, 1, f"EVS100MI.ImportFile {filename}")
        with requests.Session() as session:
            ok, message = MIG_Api.process_file_in_m3(dest, filename, session)
        progress(1, 1, message)
        if not ok:
            raise MIG_Api.M3ApiError(message)
        return {"ok": ok, "message": message, "file": filename}

    return jsonify(status="success", job_id=start_job("Sorting Options process", work))


# =============================================================================
# Sorting Orders  (SOURCE + DEST, optional PGNM filter, EVS100 file)
# =============================================================================

@app.route("/api/sorting-orders/fetch-diff", methods=["POST"])
def api_sorting_orders_fetch_diff():
    data = request.get_json(force=True) or {}
    source = _tenant(data.get("source_tenant_id"))
    dest = _tenant(data.get("dest_tenant_id"))
    pgnm_filter = (data.get("pgnm_filter") or "").strip().upper()

    def work(progress):
        progress(0, 2, "CRS022MI.LstSortOrder (SOURCE)")
        with requests.Session() as session:
            source_all = MIG_SyncSortingOrders.list_all_sort_orders(source, session, "SOURCE", pgnm_filter)
        progress(1, 2, "CRS022MI.LstSortOrder (DEST)")
        with requests.Session() as session:
            dest_all = MIG_SyncSortingOrders.list_all_sort_orders(dest, session, "DEST", pgnm_filter)
        progress(2, 2, "Comparing SOURCE vs DEST")
        result = MIG_SyncSortingOrders.diff_sort_orders(source_all, dest_all)
        result["pgnm_filter"] = pgnm_filter
        return result

    return jsonify(status="success", job_id=start_job("Sorting Orders fetch+diff", work))


@app.route("/api/sorting-orders/export", methods=["POST"])
def api_sorting_orders_export():
    data = request.get_json(force=True) or {}
    result = _job_result(data.get("job_id"))
    add_records = [MIG_SyncSortingOrders.build_payload_record(r) for r in result["to_add"]]
    chg_records = [MIG_SyncSortingOrders.build_payload_record(src) for src, _ in result["to_chg"]]
    path = MIG_SyncSortingOrders.export_evs100_xlsx(add_records, chg_records, EVS100_TO_PROCESS)
    return jsonify(status="success", file=path.name, download="evs100/" + path.name,
                   counts={"add": len(add_records), "change": len(chg_records)},
                   message=f"{len(add_records)} add + {len(chg_records)} change record(s) exported to {path.name}.")


@app.route("/api/sorting-orders/upload", methods=["POST"])
def api_sorting_orders_upload():
    data = request.get_json(force=True) or {}
    dest = _tenant(data.get("dest_tenant_id"))
    filename = Path((data.get("filename") or "")).name
    path = EVS100_TO_PROCESS / filename
    if not path.is_file():
        return jsonify(status="error", message=f"{filename} not found."), 404

    def work(progress):
        progress(0, 1, f"Uploading {filename}")
        with requests.Session() as session:
            ok, message = MIG_Api.upload_file_to_m3(dest, path, session)
        progress(1, 1, message)
        if not ok:
            raise MIG_Api.M3ApiError(message)
        return {"ok": ok, "message": message, "file": filename}

    return jsonify(status="success", job_id=start_job("Sorting Orders upload", work))


@app.route("/api/sorting-orders/process-in-m3", methods=["POST"])
def api_sorting_orders_process():
    data = request.get_json(force=True) or {}
    dest = _tenant(data.get("dest_tenant_id"))
    filename = Path((data.get("filename") or "")).name

    def work(progress):
        progress(0, 1, f"EVS100MI.ImportFile {filename}")
        with requests.Session() as session:
            ok, message = MIG_Api.process_file_in_m3(dest, filename, session)
        progress(1, 1, message)
        if not ok:
            raise MIG_Api.M3ApiError(message)
        return {"ok": ok, "message": message, "file": filename}

    return jsonify(status="success", job_id=start_job("Sorting Orders process", work))


# =============================================================================
# Transl Data  (SOURCE + DEST, both global_scope=True, EVS100 file + direct API)
# =============================================================================

@app.route("/api/transl-data/fetch-diff", methods=["POST"])
def api_transl_data_fetch_diff():
    data = request.get_json(force=True) or {}
    source = _tenant(data.get("source_tenant_id"))

    def work(progress):
        progress(0, 3, "EXPORTMI.Select MBMTRN (header)")
        with requests.Session() as session:
            hdr_rows = MIG_SyncTranslData.fetch_exportmi(
                source, session, MIG_SyncTranslData.EXPORTMI_HDR_QUERY, "MBMTRN")
            progress(1, 3, "EXPORTMI.Select MBMTRD (details)")
            dtl_rows = MIG_SyncTranslData.fetch_exportmi(
                source, session, MIG_SyncTranslData.EXPORTMI_DTL_QUERY, "MBMTRD")
        progress(2, 3, "Joining MBMTRN + MBMTRD on IDTR")
        trn_records, trd_records = MIG_SyncTranslData.build_records(hdr_rows, dtl_rows, source.company)
        progress(3, 3, "Done")
        return {
            "hdr_rows": hdr_rows, "dtl_rows": dtl_rows,
            "trn_records": trn_records, "trd_records": trd_records,
            "preview": trd_records[:50],
            "counts": {"headers": len(hdr_rows), "details": len(dtl_rows),
                       "trn_records": len(trn_records), "trd_records": len(trd_records)},
        }

    return jsonify(status="success", job_id=start_job("Transl Data fetch+diff", work))


@app.route("/api/transl-data/export", methods=["POST"])
def api_transl_data_export():
    data = request.get_json(force=True) or {}
    result = _job_result(data.get("job_id"))
    path = MIG_SyncTranslData.export_evs100_xlsx(result["trn_records"], result["trd_records"], EVS100_TO_PROCESS)
    return jsonify(status="success", file=path.name, download="evs100/" + path.name,
                   message=f"{len(result['trd_records'])} detail record(s) exported to {path.name}.")


@app.route("/api/transl-data/push-api", methods=["POST"])
def api_transl_data_push_api():
    """Direct CRS881MI.AddTranslation / AddTranslData calls against DEST."""
    data = request.get_json(force=True) or {}
    dest = _tenant(data.get("dest_tenant_id"))
    prior = _job_result(data.get("job_id"))
    batch_size = MIG_SyncTranslData.DEFAULT_BATCH_SIZE
    trn_records = prior["trn_records"]
    trd_records_dest = [{**rec, "CONO": dest.company} for rec in prior["trd_records"]]

    def work(progress):
        with requests.Session() as session:
            trn_successes, trn_failures = MIG_SyncTranslData.run_add_translation_batched(
                dest, trn_records, session, batch_size, progress=progress)
            trd_successes, trd_failures = MIG_SyncTranslData.run_add_transl_data_batched(
                dest, trd_records_dest, session, batch_size, progress=progress)

        result = {
            "counts": {"trn_success": len(trn_successes), "trn_failed": len(trn_failures),
                       "trd_success": len(trd_successes), "trd_failed": len(trd_failures)},
        }
        if trn_failures or trd_failures:
            path = MIG_SyncTranslData.export_api_errors_xlsx(
                trn_successes, trn_failures, trd_successes, trd_failures,
                OUTPUT_DIR / "transl_data")
            result["error_file"] = path.name
            result["download"] = "transl_data/" + path.name
        return result

    return jsonify(status="success", job_id=start_job("Transl Data push API", work))


@app.route("/api/transl-data/upload", methods=["POST"])
def api_transl_data_upload():
    data = request.get_json(force=True) or {}
    dest = _tenant(data.get("dest_tenant_id"))
    filename = Path((data.get("filename") or "")).name
    path = EVS100_TO_PROCESS / filename
    if not path.is_file():
        return jsonify(status="error", message=f"{filename} not found."), 404

    def work(progress):
        progress(0, 1, f"Uploading {filename}")
        with requests.Session() as session:
            ok, message = MIG_Api.upload_file_to_m3(dest, path, session)
        progress(1, 1, message)
        if not ok:
            raise MIG_Api.M3ApiError(message)
        return {"ok": ok, "message": message, "file": filename}

    return jsonify(status="success", job_id=start_job("Transl Data upload", work))


@app.route("/api/transl-data/process-in-m3", methods=["POST"])
def api_transl_data_process():
    data = request.get_json(force=True) or {}
    dest = _tenant(data.get("dest_tenant_id"))
    filename = Path((data.get("filename") or "")).name

    def work(progress):
        progress(0, 1, f"EVS100MI.ImportFile {filename}")
        with requests.Session() as session:
            ok, message = MIG_Api.process_file_in_m3(dest, filename, session)
        progress(1, 1, message)
        if not ok:
            raise MIG_Api.M3ApiError(message)
        return {"ok": ok, "message": message, "file": filename}

    return jsonify(status="success", job_id=start_job("Transl Data process", work))


# =============================================================================
# Field Groups  (single tenant, long-running MNS320 poll)
# =============================================================================

@app.route("/api/field-groups/run", methods=["POST"])
def api_field_groups_run():
    data = request.get_json(force=True) or {}
    tenant = _tenant(data.get("tenant_id"))
    poll_interval = int(data.get("poll_interval") or MIG_GenerateFieldGroups.POLL_INTERVAL)

    def work(progress):
        return MIG_GenerateFieldGroups.run_gen_standard(tenant, progress=progress, poll_interval=poll_interval)

    return jsonify(status="success", job_id=start_job("Field Groups (CMS005MI.GenStandard)", work),
                   poll_interval=poll_interval)


# =============================================================================
# Export Summary  (no tenant - pure .log -> .xlsx converter + sqlite lookup)
# =============================================================================

_EXPORT_SUMMARY_UPLOAD_DIR = UPLOAD_DIR / "export_summary"
_EXPORT_SUMMARY_OUTPUT_DIR = MIG_ExportSummary.DEFAULT_OUTPUT_FOLDER


def _process_export_summary_logs(paths: list[Path]) -> list[dict]:
    _EXPORT_SUMMARY_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    results = []
    for p in paths:
        res = MIG_ExportSummary.process_log(p, out_dir=_EXPORT_SUMMARY_OUTPUT_DIR)
        if res.get("ok"):
            res["download"] = "export_summary/" + res["file"]
        results.append(res)
    return results


@app.route("/api/export-summary/upload", methods=["POST"])
def api_export_summary_upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify(status="error", message="No files received."), 400

    _EXPORT_SUMMARY_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    saved = []
    errors = []
    for fs in files:
        name = Path(fs.filename or "").name
        if not name.lower().endswith(".log"):
            errors.append({"file": name, "message": "Not a .log file."})
            continue
        dest = _EXPORT_SUMMARY_UPLOAD_DIR / name
        fs.save(dest)
        saved.append(dest)

    results = _process_export_summary_logs(saved)
    return jsonify(status="success", results=results, errors=errors)


@app.route("/api/export-summary/scan", methods=["POST"])
def api_export_summary_scan():
    """Process whatever .log files are already sitting in the upload dir."""
    _EXPORT_SUMMARY_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    paths = sorted(_EXPORT_SUMMARY_UPLOAD_DIR.glob("*.log"))
    if not paths:
        return jsonify(status="success", results=[],
                       message=f"No .log files found in {_EXPORT_SUMMARY_UPLOAD_DIR}.")
    return jsonify(status="success", results=_process_export_summary_logs(paths))


# =============================================================================
# Migration Review  (single tenant, upload/scan .xlsx workbooks)
# =============================================================================

@app.route("/api/migration-review/upload", methods=["POST"])
def api_migration_review_upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify(status="error", message="No files received."), 400

    MIG_MigrationReview.INPUT_DIR.mkdir(parents=True, exist_ok=True)
    saved = []
    errors = []
    for fs in files:
        name = Path(fs.filename or "").name
        if not name.lower().endswith(".xlsx") or name.startswith("~$"):
            errors.append({"file": name, "message": "Not a .xlsx file."})
            continue
        dest = MIG_MigrationReview.INPUT_DIR / name
        fs.save(dest)
        saved.append(name)

    return jsonify(status="success", saved=saved, errors=errors,
                   message=f"{len(saved)} workbook(s) uploaded; run Scan to process them.")


@app.route("/api/migration-review/scan", methods=["POST"])
def api_migration_review_scan():
    data = request.get_json(force=True) or {}
    tenant = _tenant(data.get("tenant_id"))
    files = MIG_MigrationReview.list_pending_workbooks()
    if not files:
        return jsonify(status="success", job_id=None, results=[],
                       message=f"No workbooks found in {MIG_MigrationReview.INPUT_DIR}.")

    def work(progress):
        results = []
        for i, path in enumerate(files):
            progress(i, len(files), f"Processing {path.name}")
            res = MIG_MigrationReview.process_workbook(path, tenant, progress=progress)
            if res.get("ok"):
                res["download"] = "migration_review/" + res["output_file"]
            results.append(res)
        progress(len(files), len(files), "Done")
        return {"results": results}

    return jsonify(status="success", job_id=start_job("Migration Review scan", work))


# =============================================================================
# Entry point
# =============================================================================

def main() -> None:
    ap = argparse.ArgumentParser(description="MIG Sync front end.")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=5059)
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO,
                         format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    for d in (UPLOAD_DIR, OUTPUT_DIR, _EXPORT_SUMMARY_UPLOAD_DIR, _EXPORT_SUMMARY_OUTPUT_DIR,
              MIG_MigrationReview.INPUT_DIR, MIG_MigrationReview.PROCESSED_DIR,
              MIG_MigrationReview.OUTPUT_DIR):
        d.mkdir(parents=True, exist_ok=True)

    log.info("ionapi   : %s", IONAPI_DIR)
    log.info("evs100   : %s", EVS100_TO_PROCESS)
    log.info("uploads  : %s", UPLOAD_DIR)
    log.info("output   : %s", OUTPUT_DIR)

    app.run(host=args.host, port=args.port, debug=args.debug, threaded=True)


if __name__ == "__main__":
    main()
