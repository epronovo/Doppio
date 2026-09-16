# MIG_GenerateFieldGroups.py
"""
MIG_GenerateFieldGroups.py
----------------------------
Runs CMS005MI.GenStandard (ST02=1) against a single tenant and waits for the
background M3 job to complete.

Steps
-----
1.  Authenticate (a single tenant - handled by MIG_App's /api/tenant/connect).
2.  Call CMS005MI.GenStandard -> capture the returned BJNO.
3.  Poll MNS320MI.Get every POLL_INTERVAL seconds:
      - No record returned  -> job completed normally.
      - STAT != waiting     -> job ended in error; return the full result.

Ported to a library used by MIG_App.py: the original polling loop's
time.sleep(POLL_INTERVAL) + print() becomes a job progress(done, total,
message) callback that the browser polls via /api/job/<id>, and the final
outcome is returned as a result dict instead of just printed.
"""

from __future__ import annotations

import time

import requests

import MIG_Api

# =============================================================================
# Constants
# =============================================================================

# How long to wait between job-status polls (seconds). Matches the MNS320
# polling cadence used by the original CLI script - keep this the real
# default, not a shortened one, or a caller will hammer M3.
POLL_INTERVAL = 60

# MNS320MI STAT values that mean "keep waiting". 00 = not yet started,
# 20 = running.
STATUS_WAIT = frozenset({"00", "20"})

JOB_STATUS_FIELDS = [
    "BJNO", "STAT", "JOBQ", "JBPR", "JNA", "JNU",
    "CONO", "DIVI", "RGDT", "RGTM", "LMDT", "LMTM", "CHID",
]


def _noop_progress(done, total, message=""):
    pass


def norm(v) -> str:
    return "" if v is None else str(v).strip()


# =============================================================================
# Step 2 - GenStandard
# =============================================================================

def gen_standard(tenant: MIG_Api.Tenant, session: requests.Session) -> str:
    """
    Calls CMS005MI.GenStandard with ST02=1.
    Returns the BJNO string, or raises MIG_Api.M3ApiError on failure.
    """
    payload = {
        "program": "CMS005MI",
        "transactions": [{"transaction": "GenStandard", "record": {"ST02": "1"},
                           "selectedColumns": ["BJNO"]}],
    }
    data = MIG_Api.post_to_m3(tenant, payload, session)

    for result in data.get("results", []):
        err = (result.get("errorMessage") or "").strip()
        if err:
            raise MIG_Api.M3ApiError(f"CMS005MI.GenStandard error: {err}")
        for record in result.get("records", []):
            bjno = (record.get("BJNO") or "").strip()
            if bjno:
                return bjno

    raise MIG_Api.M3ApiError("CMS005MI.GenStandard did not return a BJNO.")


# =============================================================================
# Step 3 - Poll MNS320MI.Get until job completes
# =============================================================================

def poll_job_status(tenant: MIG_Api.Tenant, bjno: str,
                     progress=_noop_progress,
                     poll_interval: int = POLL_INTERVAL,
                     sleep_fn=time.sleep) -> dict:
    """
    Polls MNS320MI.Get for the given BJNO every poll_interval seconds. A fresh
    HTTP session is opened for each attempt to avoid stale connections.

    Termination:
      - No records returned      -> completed normally.
      - STAT not in STATUS_WAIT  -> ended in error; full record is returned.

    Returns a result dict describing how the job ended:
        {"outcome": "completed", "bjno": bjno, "attempts": n}
      or
        {"outcome": "error", "bjno": bjno, "attempts": n, "stat": "...",
         "record": {...}}
    progress(done, total, message) is called once per poll attempt (total is
    always 0 - this is an open-ended wait, not a boundable count) so the
    caller's job/<id> poll route has something to show.
    """
    attempt = 0
    while True:
        attempt += 1
        progress(attempt, 0, f"Waiting for BJNO={bjno} (poll #{attempt})")

        payload = {
            "program": "MNS320MI",
            "transactions": [{"transaction": "Get", "record": {"BJNO": bjno},
                               "selectedColumns": JOB_STATUS_FIELDS}],
        }
        try:
            with requests.Session() as poll_session:
                data = MIG_Api.post_to_m3(tenant, payload, poll_session)
        except MIG_Api.M3ApiError as exc:
            progress(attempt, 0, f"Poll attempt {attempt} failed: {exc}; retrying in {poll_interval}s")
            sleep_fn(poll_interval)
            continue

        records: list[dict] = []
        for result in data.get("results", []):
            for record in result.get("records", []):
                records.append(record)

        if not records:
            progress(attempt, 0, f"Job BJNO={bjno} completed normally.")
            return {"outcome": "completed", "bjno": bjno, "attempts": attempt}

        rec = records[0]
        stat = norm(rec.get("STAT", ""))

        if stat not in STATUS_WAIT:
            progress(attempt, 0, f"Job BJNO={bjno} ended with STAT={stat}")
            return {"outcome": "error", "bjno": bjno, "attempts": attempt,
                    "stat": stat,
                    "record": {f: norm(rec.get(f, "")) for f in JOB_STATUS_FIELDS if norm(rec.get(f, ""))}}

        progress(attempt, 0,
                 f"STAT={stat} JBPR={norm(rec.get('JBPR',''))} JOBQ={norm(rec.get('JOBQ',''))} - "
                 f"next check in {poll_interval}s")
        sleep_fn(poll_interval)


def run_gen_standard(tenant: MIG_Api.Tenant, progress=_noop_progress,
                      poll_interval: int = POLL_INTERVAL) -> dict:
    """
    Full flow: submit CMS005MI.GenStandard, then poll until it ends. Returns
    {"bjno": ..., "elapsed_seconds": ..., **poll_job_status result}.
    """
    start_time = time.time()
    with requests.Session() as session:
        bjno = gen_standard(tenant, session)

    progress(0, 0, f"Job submitted. BJNO={bjno}")
    result = poll_job_status(tenant, bjno, progress=progress, poll_interval=poll_interval)
    result["elapsed_seconds"] = round(time.time() - start_time, 1)
    return result
