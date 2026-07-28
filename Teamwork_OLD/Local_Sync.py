"""
Local_Sync.py
-------------
SQLite-only orchestrator for local development.
Runs Local_Contracts, Local_TimeData, and Local_StatementDataMonthly in sequence.

Usage:
  python3 Local_Sync.py
  python3 Local_Sync.py --reporting-date YYYY-MM-DD

If --reporting-date is omitted, the most recent Saturday is used.
"""

import sys
from datetime import date, datetime, timedelta

import Local_Contracts
import Local_TimeData
import Local_StatementDataMonthly


def _last_saturday() -> str:
    today = date.today()
    days_since_saturday = (today.weekday() - 5) % 7
    return (today - timedelta(days=days_since_saturday)).isoformat()


def start_function(reporting_date: str = None):
    if reporting_date is None:
        reporting_date = _last_saturday()

    start_timestamp = datetime.now()
    print(f"=== Local Sync Started: {start_timestamp} ===")
    print(f"Backend: SQLite")
    print(f"Reporting Date: {reporting_date}\n")

    print("=" * 70)
    print("  Local_Contracts")
    print("=" * 70)
    Local_Contracts.start_function()

    print()
    print("=" * 70)
    print("  Local_TimeData")
    print("=" * 70)
    Local_TimeData.start_function()

    print()
    print("=" * 70)
    print("  Local_StatementDataMonthly")
    print("=" * 70)
    Local_StatementDataMonthly.start_function(reporting_date=reporting_date)

    end_timestamp = datetime.now()
    print(f"\n=== Local Sync Finished: {end_timestamp} ===")
    print(f"Total Duration (in seconds): {(end_timestamp - start_timestamp).total_seconds()}")


if __name__ == "__main__":
    rd = None

    if '--reporting-date' in sys.argv:
        idx = sys.argv.index('--reporting-date')
        if idx + 1 < len(sys.argv):
            rd = sys.argv[idx + 1]

    start_function(reporting_date=rd)
