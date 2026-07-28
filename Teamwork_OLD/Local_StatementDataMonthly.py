"""
Local_StatementDataMonthly.py
------------------------------
SQLite-only version of Teamwork_StatementDataMonthly.py for local development.
Reads from and writes to teamwork.db — no BigQuery dependency.

Usage:
  python Local_StatementDataMonthly.py YYYY-MM-DD
"""

import sqlite3
import sys
from datetime import date, datetime, timedelta

import pandas as pd

# ── Configuration ─────────────────────────────────────────────────────────────
DB_FILE    = 'teamwork.db'
TABLE_NAME = 'StatementDataMonthly'


# ── SQLite query ──────────────────────────────────────────────────────────────
SQLITE_QUERY = """
WITH RECURSIVE months AS (
    SELECT
        project_id,
        project_name,
        company_name,
        section,
        record_name,
        Start_Date    AS quarter_start,
        End_Date      AS quarter_end,
        Additional, Hours, Overage, Rate, URL,
        Start_Date    AS month_start,
        min(date(Start_Date, 'start of month', '+1 month', '-1 day'), End_Date) AS month_end,
        (strftime('%Y', End_Date) - strftime('%Y', Start_Date)) * 12
            + (strftime('%m', End_Date) - strftime('%m', Start_Date)) + 1 AS num_months
    FROM ContractFieldValues
    WHERE Start_Date IS NOT NULL AND Start_Date != ''
    UNION ALL
    SELECT
        project_id, project_name, company_name, section, record_name,
        quarter_start, quarter_end,
        Additional, Hours, Overage, Rate, URL,
        date(month_start, 'start of month', '+1 month') AS month_start,
        min(date(month_start, 'start of month', '+2 months', '-1 day'), quarter_end) AS month_end,
        num_months
    FROM months
    WHERE date(month_start, 'start of month', '+1 month') <= quarter_end
),
base AS (
    SELECT
        m.project_id,
        m.project_name,
        m.company_name,
        m.section,
        m.record_name,
        m.month_start,
        m.month_end,
        CASE WHEN m.month_start = m.quarter_start THEN m.Additional END AS Additional,
        CASE WHEN m.month_start = m.quarter_start THEN m.Hours END AS Hours,
        CASE WHEN m.month_start = m.quarter_start THEN m.Overage END AS Overage,
        m.Rate,
        ROUND(CAST(m.Hours AS REAL) / m.num_months, 4) AS purchased_hours,
        CASE WHEN m.month_start = m.quarter_start
             THEN COALESCE(CAST(m.Additional AS REAL), 0)
             ELSE 0
        END AS additional_mth,
        CASE WHEN m.month_start = m.quarter_start
             THEN COALESCE(CAST(m.Overage AS REAL), 0)
             ELSE 0
        END AS overage_mth,
        COALESCE(SUM(te.Billable_Hours), 0) AS utilized_hours,
        COALESCE(SUM(CASE WHEN te.Date <= :reporting_date THEN te.Billable_Hours ELSE 0 END), 0) AS utilized_hours_to_date
    FROM months m
    LEFT JOIN TimeData te
        ON  te.Project_ID = m.project_id
        AND te.Date >= m.month_start
        AND te.Date <= m.month_end
    GROUP BY m.project_id, m.project_name, m.company_name, m.section, m.record_name, m.month_start
)
SELECT
    project_id AS "Project_ID",
    section,
    company_name AS "Customer_Name",
    project_name AS "Project_Name",
    record_name AS "Contract_Period_Name",
    month_start AS "Period_Start_Date",
    month_end AS "Period_End_Date",
    :reporting_date AS "Reporting_Date",
    Rate AS "Rate_per_hour",
    purchased_hours AS "Purchased_Hours",
    CASE WHEN Hours IS NOT NULL THEN ROUND(COALESCE(SUM((purchased_hours + additional_mth + overage_mth) - utilized_hours) OVER w_prev, 0),2) ELSE 0 END AS "Approved_rollover_Qtr",
    ROUND(COALESCE(SUM((purchased_hours + additional_mth + overage_mth) - utilized_hours) OVER w_prev, 0),2) AS "Approved_rollover_Mth",
    COALESCE(ROUND(Hours + Additional + COALESCE(SUM((purchased_hours + additional_mth + overage_mth) - utilized_hours) OVER w_prev, 0),0),2) AS "Purchased_Hours_Qtr",
    ROUND((purchased_hours + additional_mth + overage_mth) + COALESCE(SUM((purchased_hours + additional_mth + overage_mth) - utilized_hours) OVER w_prev, 0),2) AS "Puchased_Hours_Mth",
    ROUND(utilized_hours, 2) AS "Utilized_Hours",
    ROUND(utilized_hours_to_date, 2) AS "Utilized_Hours_to_Reporting_Date",
    ROUND(COALESCE(SUM((purchased_hours + additional_mth + overage_mth) - utilized_hours) OVER w_curr, 0),2) AS "Remaining_Hours",
    CASE
        WHEN date(:reporting_date) >= date(month_end)   THEN 1.0
        WHEN date(:reporting_date) <= date(month_start) THEN 0.0
        ELSE ROUND(
            (julianday(:reporting_date) - julianday(month_start)) /
            (julianday(month_end)       - julianday(month_start)), 4)
    END AS "Schedule_Elapsed_%",
    CASE WHEN Overage <> '' THEN Overage ELSE 0 END AS "Overage_Hours",
    CASE WHEN Additional <> '' THEN Additional ELSE 0 END AS "Addition_Pre_purchased_hours",
    COALESCE(Hours,0) AS "Total_Pre_purchased_hours_(QTR)",
    CASE WHEN Hours IS NOT NULL AND record_name = 'Q1' THEN COALESCE(Hours,0) + ROUND(COALESCE(SUM((purchased_hours + additional_mth + overage_mth) - utilized_hours) OVER w_prev, 0),2) ELSE COALESCE(Hours,0) END AS "Starting_Hours_Qtr"
FROM base
WINDOW
    w_prev AS (PARTITION BY project_id ORDER BY month_start ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING),
    w_curr AS (PARTITION BY project_id ORDER BY month_start ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)
ORDER BY project_id, section, month_start
"""


# ── Helpers ───────────────────────────────────────────────────────────────────

def _last_saturday() -> str:
    today = date.today()
    days_since_saturday = (today.weekday() - 5) % 7
    return (today - timedelta(days=days_since_saturday)).isoformat()


# ── Entry point ───────────────────────────────────────────────────────────────

def start_function(reporting_date: str = None):
    if not reporting_date:
        reporting_date = _last_saturday()

    start_timestamp = datetime.now()
    print(f"Script Started: {start_timestamp}")
    print(f"Reporting Date: {reporting_date}")

    with sqlite3.connect(DB_FILE) as conn:
        df = pd.read_sql(SQLITE_QUERY, conn, params={'reporting_date': reporting_date})
        df.to_sql(TABLE_NAME, conn, if_exists='replace', index=False)
    print(f"Written {len(df)} rows to {DB_FILE} → {TABLE_NAME}")

    end_timestamp = datetime.now()
    print(f"Script Finished: {end_timestamp}")
    print(f"Run Duration (in seconds): {(end_timestamp - start_timestamp).total_seconds()}")


if __name__ == "__main__":
    rd = next((a for a in sys.argv[1:] if not a.startswith('--')), None)
    if not rd:
        print("Usage: python3 Local_StatementDataMonthly.py YYYY-MM-DD")
        sys.exit(1)
    start_function(reporting_date=rd)
