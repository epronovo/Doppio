"""
Teamwork_StatementDataMonthly.py
---------------------------------
Expands AMS contract quarters into months, joins against logged time,
and writes the result to BigQuery.StatementDataMonthly.

Requires a reporting_date (YYYY-MM-DD) — all "to date" utilisation columns
and the Schedule Elapsed % are computed relative to that date.

Usage:
  python Teamwork_StatementDataMonthly.py YYYY-MM-DD
"""

import sys
from datetime import date, datetime, timedelta

# ── Configuration ─────────────────────────────────────────────────────────────
TABLE_NAME = 'StatementDataMonthly'

BQ_PROJECT = 'spatial-earth-492100-b7'
BQ_DATASET = 'teamwork'
KEY_PATH   = '/Users/ericpronovost/Downloads/spatial-earth-492100-b7-cd5d8fb255b4.json'


# ── BigQuery query ─────────────────────────────────────────────────────────────
BQ_QUERY = f"""
WITH RECURSIVE
cfv AS (
    SELECT *
    FROM `{BQ_PROJECT}.{BQ_DATASET}.ContractFieldValues`
    WHERE Start_Date IS NOT NULL
      AND End_Date   IS NOT NULL
),
months AS (
    SELECT
        project_id,
        project_name,
        company_name,
        section,
        record_name,
        DATE(Start_Date)  AS quarter_start,
        DATE(End_Date)    AS quarter_end,
        Additional, Hours, Overage, Rate, URL,
        DATE(Start_Date)  AS month_start,
        LEAST(
            DATE_SUB(DATE_ADD(DATE_TRUNC(DATE(Start_Date), MONTH), INTERVAL 1 MONTH), INTERVAL 1 DAY),
            DATE(End_Date)
        ) AS month_end,
        DATE_DIFF(DATE(End_Date), DATE(Start_Date), MONTH) + 1 AS num_months
    FROM cfv
    UNION ALL
    SELECT
        project_id, project_name, company_name, section, record_name,
        quarter_start, quarter_end,
        Additional, Hours, Overage, Rate, URL,
        DATE_ADD(DATE_TRUNC(month_start, MONTH), INTERVAL 1 MONTH) AS month_start,
        LEAST(
            DATE_SUB(DATE_ADD(DATE_TRUNC(month_start, MONTH), INTERVAL 2 MONTH), INTERVAL 1 DAY),
            quarter_end
        ) AS month_end,
        num_months
    FROM months
    WHERE DATE_ADD(DATE_TRUNC(month_start, MONTH), INTERVAL 1 MONTH) <= quarter_end
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
        CASE WHEN m.month_start = m.quarter_start THEN m.Hours END     AS Hours,
        CASE WHEN m.month_start = m.quarter_start THEN m.Overage END   AS Overage,
        m.Rate,
        ROUND(SAFE_CAST(m.Hours AS FLOAT64) / m.num_months, 4) AS purchased_hours,
        CASE WHEN m.month_start = m.quarter_start
             THEN COALESCE(SAFE_CAST(m.Additional AS FLOAT64), 0)
             ELSE 0
        END AS additional_mth,
        CASE WHEN m.month_start = m.quarter_start
             THEN COALESCE(SAFE_CAST(m.Overage AS FLOAT64), 0)
             ELSE 0
        END AS overage_mth,
        COALESCE(SUM(te.Billable_Hours), 0) AS utilized_hours,
        COALESCE(SUM(CASE WHEN DATE(te.Date) <= @reporting_date THEN te.Billable_Hours ELSE 0 END), 0) AS utilized_hours_to_date
    FROM months m
    LEFT JOIN `{BQ_PROJECT}.{BQ_DATASET}.TimeData` te
        ON  te.Project_ID  = m.project_id
        AND DATE(te.Date) >= m.month_start
        AND DATE(te.Date) <= m.month_end
    GROUP BY
        m.project_id, m.project_name, m.company_name, m.section, m.record_name,
        m.month_start, m.month_end, m.quarter_start,
        m.Additional, m.Hours, m.Overage, m.Rate, m.num_months
)
SELECT
    project_id                                                                                                                                                                          AS `Project_ID`,
    section,
    company_name                                                                                                                                                                        AS `Customer_Name`,
    project_name                                                                                                                                                                        AS `Project_Name`,
    record_name                                                                                                                                                                         AS `Contract_Period_Name`,
    month_start                                                                                                                                                                         AS `Period_Start_Date`,
    month_end                                                                                                                                                                           AS `Period_End_Date`,
    @reporting_date                                                                                                                                                                     AS `Reporting_Date`,
    SAFE_CAST(Rate AS FLOAT64)                                                                                                                                                          AS `Rate_per_hour`,
    purchased_hours                                                                                                                                                                     AS `Purchased_Hours`,
    CASE WHEN Hours IS NOT NULL THEN ROUND(COALESCE(SUM((purchased_hours + additional_mth + overage_mth) - utilized_hours) OVER w_prev, 0),2) ELSE 0.0 END                            AS `Approved_rollover_Qtr`,
    ROUND(COALESCE(SUM((purchased_hours + additional_mth + overage_mth) - utilized_hours) OVER w_prev, 0), 2)                                                                         AS `Approved_rollover_Mth`,
    COALESCE(SAFE_CAST(Hours AS FLOAT64) + COALESCE(SAFE_CAST(Additional AS FLOAT64), 0) + ROUND(COALESCE(SUM((purchased_hours + additional_mth + overage_mth) - utilized_hours) OVER w_prev, 0),2), 0) AS `Purchased_Hours_Qtr`,
    ROUND((purchased_hours + additional_mth + overage_mth) + COALESCE(SUM((purchased_hours + additional_mth + overage_mth) - utilized_hours) OVER w_prev, 0), 2)                     AS `Puchased_Hours_Mth`,
    ROUND(utilized_hours, 2)                                                                                                                                                            AS `Utilized_Hours`,
    ROUND(utilized_hours_to_date, 2)                                                                                                                                                    AS `Utilized_Hours_to_Reporting_Date`,
    ROUND(COALESCE(SUM((purchased_hours + additional_mth + overage_mth) - utilized_hours) OVER w_curr, 0), 2)                                                                         AS `Remaining_Hours`,
    CASE
        WHEN @reporting_date >= month_end   THEN 1.0
        WHEN @reporting_date <= month_start THEN 0.0
        ELSE ROUND(SAFE_DIVIDE(
            DATE_DIFF(@reporting_date, month_start, DAY),
            DATE_DIFF(month_end,       month_start, DAY)), 4)
    END                                                                                                                                                                                 AS `Schedule_Elapsed_Pct`,
    CASE WHEN Overage != '' THEN SAFE_CAST(Overage AS FLOAT64) ELSE 0.0 END                                                                                                            AS `Overage_Hours`,
    CASE WHEN Additional != '' THEN SAFE_CAST(Additional AS FLOAT64) ELSE 0.0 END                                                                                                      AS `Addition_Pre_purchased_hours`,
    COALESCE(SAFE_CAST(Hours AS FLOAT64), 0)                                                                                                                                           AS `Total_Pre_purchased_hours_QTR`,
    CASE WHEN Hours IS NOT NULL AND record_name = 'Q1' THEN COALESCE(SAFE_CAST(Hours AS FLOAT64),0) + ROUND(COALESCE(SUM((purchased_hours + additional_mth + overage_mth) - utilized_hours) OVER w_prev, 0),2) ELSE COALESCE(SAFE_CAST(Hours AS FLOAT64),0) END AS `Starting_Hours_Qtr`
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


def get_bq_client():
    from google.cloud import bigquery
    from google.oauth2 import service_account
    creds = service_account.Credentials.from_service_account_file(
        KEY_PATH,
        scopes=["https://www.googleapis.com/auth/bigquery"],
    )
    return bigquery.Client(project=BQ_PROJECT, credentials=creds)


# ── Entry point ───────────────────────────────────────────────────────────────

def start_function(reporting_date: str = None):
    if not reporting_date:
        reporting_date = _last_saturday()

    start_timestamp = datetime.now()
    print(f"Script Started: {start_timestamp}")
    print(f"Reporting Date: {reporting_date}")

    from google.cloud import bigquery
    client = get_bq_client()
    table_id = f"{BQ_PROJECT}.{BQ_DATASET}.{TABLE_NAME}"
    job_config = bigquery.QueryJobConfig(
        destination=table_id,
        write_disposition=bigquery.WriteDisposition.WRITE_TRUNCATE,
        query_parameters=[
            bigquery.ScalarQueryParameter("reporting_date", "DATE", reporting_date),
        ],
    )
    client.query(BQ_QUERY, job_config=job_config).result()
    print(f"Written to BigQuery {table_id}")

    end_timestamp = datetime.now()
    print(f"Script Finished: {end_timestamp}")
    print(f"Run Duration (in seconds): {(end_timestamp - start_timestamp).total_seconds()}")


if __name__ == "__main__":
    rd = next((a for a in sys.argv[1:] if not a.startswith('--')), None)
    if not rd:
        print("Usage: python3 Teamwork_StatementDataMonthly.py YYYY-MM-DD")
        sys.exit(1)
    start_function(reporting_date=rd)
