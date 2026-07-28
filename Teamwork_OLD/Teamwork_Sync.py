"""
Teamwork_Sync.py
----------------
Orchestrator that runs Teamwork_Contracts.py, Teamwork_TimeData.py, and
Teamwork_StatementDataMonthly.py against BigQuery.

Cloud Run entry point: POST /trigger
  Body (optional JSON): {"reporting_date": "YYYY-MM-DD"}

Local usage:
  python3 Teamwork_Sync.py
  python3 Teamwork_Sync.py --reporting-date YYYY-MM-DD

If reporting_date is omitted, the most recent Saturday is used.
"""

import sys
from datetime import date, datetime, timedelta

import Teamwork_Contracts
import Teamwork_TimeData
import Teamwork_StatementDataMonthly


def _last_saturday() -> str:
    today = date.today()
    days_since_saturday = (today.weekday() - 5) % 7
    return (today - timedelta(days=days_since_saturday)).isoformat()


def start_function(reporting_date: str = None):
    if reporting_date is None:
        reporting_date = _last_saturday()

    start_timestamp = datetime.now()
    print(f"=== Teamwork Sync Started: {start_timestamp} ===")
    print(f"Backend: BigQuery")
    print(f"Reporting Date: {reporting_date}\n")

    print("=" * 70)
    print("  Teamwork_Contracts")
    print("=" * 70)
    Teamwork_Contracts.start_function()

    print()
    print("=" * 70)
    print("  Teamwork_TimeData")
    print("=" * 70)
    Teamwork_TimeData.start_function()

    print()
    print("=" * 70)
    print("  Teamwork_StatementDataMonthly")
    print("=" * 70)
    Teamwork_StatementDataMonthly.start_function(reporting_date=reporting_date)

    end_timestamp = datetime.now()
    print(f"\n=== Teamwork Sync Finished: {end_timestamp} ===")
    print(f"Total Duration (in seconds): {(end_timestamp - start_timestamp).total_seconds()}")


# ── Cloud Run HTTP server ──────────────────────────────────────────────────────

from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/trigger', methods=['POST'])
def trigger_sync():
    """Entry point for Cloud Scheduler."""
    body = request.get_json(silent=True)
    reporting_date = (body or {}).get('reporting_date')
    start_function(reporting_date=reporting_date)
    return jsonify({"status": "ok"}), 200


if __name__ == "__main__":
    rd = None

    if '--reporting-date' in sys.argv:
        idx = sys.argv.index('--reporting-date')
        if idx + 1 < len(sys.argv):
            rd = sys.argv[idx + 1]

    # When run directly, execute the sync without starting an HTTP server
    start_function(reporting_date=rd)
