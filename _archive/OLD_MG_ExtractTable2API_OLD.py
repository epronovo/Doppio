# MpGd_ExtractTable2API.py
import sqlite3
import subprocess
import sys
from pathlib import Path
from typing import List, Optional
from tqdm import tqdm  # pip install tqdm

DB_PATH = "/Users/ericpronovost/sqlite/doppio.db"

# ============================================================
# SQL TEMPLATE
# ============================================================
SQL_TEMPLATE = """
INSERT INTO m3Api2Table
SELECT DISTINCT {select_table}, ColumnName,
       trim(substr(pgm_panel, 1, instr(pgm_panel, '/') - 1)) AS Program,
       trim(substr(pgm_panel, instr(pgm_panel, '/') + 1)) AS Panel,
       CASE WHEN temp2.MINM IS NOT NULL THEN COALESCE(temp2.MINM,'')
            ELSE COALESCE(b.MINM,'') END AS API,
       CASE WHEN MIN(temp2.MINM) IS NOT NULL THEN COALESCE(MIN(temp2.TRNM),'')
            ELSE COALESCE(MIN(b.TRNM),'') END AS TransactionName,
       FieldName
FROM (
    SELECT TableName,ColumnName,Description,DataType,Length,Decimals,
           pgm_panel,FieldName,MINM,TRNM,Sequence
    FROM (
        SELECT mg.TableName,mg.ColumnName,mg.Description,mg.DataType,
               mg.Length,COALESCE(mg.Decimals,'') AS Decimals,
               COALESCE(CAST(MIN(c.j3pgm || ' / ' || c.j3pic1) AS TEXT),'') AS pgm_panel,
               REPLACE(SUBSTR(COALESCE(mg.ColumnName,''),-4),'LHCD','LNCD') AS FieldName,
               Sequence
        FROM m3TableCols mg
        LEFT JOIN CSEFPU c ON SUBSTR(COALESCE(c.j3fldi,''),-4) =
                             SUBSTR(COALESCE(mg.ColumnName,''),-4)
            AND c.j3pgm IN ({pgm_in})
            AND c.j3pic1 NOT IN ('A','B','C','D','P')
        WHERE mg.TableName IN ({table_in})
        GROUP BY mg.TableName,mg.ColumnName,mg.Description,
                 mg.DataType,mg.Length,mg.Decimals
        ORDER BY Sequence
    ) AS TEMP1
    LEFT JOIN cmifld a ON a.TRTP = 'I'
        AND a.MINM = '{minm}'
        AND a.FLNM = FieldName
        AND a.FLNM <> 'CHNO'
        AND a.TRNM {add_trnm_cond}
) AS temp2
LEFT JOIN cmifld b ON b.TRTP = 'I'
    AND b.MINM = '{minm}'
    AND b.FLNM = FieldName
    AND b.FLNM <> 'CHNO'
    AND b.TRNM NOT IN ({b_exclude})
    AND (b.TRNM LIKE 'Chg%' OR b.TRNM LIKE 'Upd%')
GROUP BY TableName,ColumnName,Description,DataType,Length,Decimals,
         Program,Panel,API,FieldName,Sequence
ORDER BY Sequence;
"""

# ============================================================
# HELPERS
# ============================================================
def table_exists_in_m3TableCols(conn: sqlite3.Connection, table_name: str) -> bool:
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT 1 FROM m3TableCols WHERE TableName = ? LIMIT 1",
            (table_name,)
        )
        return cur.fetchone() is not None
    except sqlite3.OperationalError:
        # m3TableCols does not exist yet
        return False


def extract_table_if_missing(table_name: str):
    script_path = Path(__file__).parent / "MpGd_ExtractTableCols.py"
    print(f"📥 Table {table_name} missing in m3TableCols — extracting...")
    subprocess.run(
        [sys.executable, str(script_path), table_name],
        check=True
    )

# ============================================================
# SQL EXECUTION
# ============================================================
def execute_sql_block(
    conn: sqlite3.Connection,
    pgm_list: List[str],
    table_list: List[str],
    minm: str,
    add_trnm: Optional[str] = None,
    b_excluded_trnms: Optional[List[str]] = None,
    table_key: Optional[str] = None,
    debug: bool = False
):
    if table_key:
        select_table = f"'CSYTAB_{table_key}' AS TableName"
    else:
        select_table = "TableName"

    table_in = ", ".join(f"'{t}'" for t in table_list)
    pgm_in = ", ".join(f"'{p}'" for p in pgm_list)
    add_trnm_cond = f"= '{add_trnm}'" if add_trnm else "LIKE 'Add%'"
    b_exclude = ", ".join(f"'{x}'" for x in b_excluded_trnms) if b_excluded_trnms else "''"

    sql = SQL_TEMPLATE.format(
        pgm_in=pgm_in,
        table_in=table_in,
        minm=minm,
        add_trnm_cond=add_trnm_cond,
        b_exclude=b_exclude,
        select_table=select_table
    )

    if debug:
        print("\n=== DEBUG SQL ===")
        print(sql[:1500])
        print("=== END DEBUG ===\n")

    conn.executescript(sql)

# ============================================================
# MAIN DRIVER
# ============================================================
def build_m3_table2api():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.executescript("""
        DROP TABLE IF EXISTS m3Api2Table;
        CREATE TABLE m3Api2Table (
            TableName TEXT,
            ColumnName TEXT,
            Program TEXT,
            Panel TEXT,
            API TEXT,
            TransactionName TEXT,
            FieldName TEXT
        );
    """)
    conn.commit()

    sql_blocks = [
        (["CRS610"], ["OCUSMA"], "CRS610MI", "Add", ["ChgAddress"], None),
        (["MMS015"], ["MITAUN"], "MMS015MI", "Add", None, None),
        (["CRS128"], ["CSYCSN"], "CRS128MI", None, None, None),
        (["MMS001"], ["MITMAS"], "MMS200MI", "AddItmBasic", ["UpdItmFac", "UpdItmWhs"], None),
        (["MMS002"], ["MITBAL"], "MMS200MI", "AddItmWhs", ["UpdItmBasic"], None),
        (["MMS003"], ["MITFAC"], "MMS200MI", "UpdItmFac", None, None),
        (["CRS620", "CRS624"], ["CIDMAS", "CIDVEN"], "CRS620MI", "AddSupplier", None, None),
        (["CRS622"], ["CIDADR"], "CRS620MI", "AddAddress", None, None),
        (["CRS692"], ["CBANAC"], "CRS692MI", "AddBankAccount2", None, None),
        (["CRS630"], ["FCHACC"], "CRS630MI", "AddAccountID", None, None),
        (["OIS002"], ["OCUSAD"], "CRS610MI", "AddAddress", None, None),
        (["MMS010"], ["MITPCE"], "MMS010MI", None, None, None),
        (["PDS001"], ["MPDHED"], "PDS001MI", None, None, None),
        (["PDS002"], ["MPDOPE"], "PDS002MI", None, None, None),
        (["PDS002"], ["MPDMAT"], "PDS002MI", None, None, None),
        (["OIS017"], ["OPRICH"], "OIS017MI", "AddPriceList", None, None),
        (["OIS021"], ["OPRICL"], "OIS017MI", "AddBasePrice",
         ["AddPriceList","ChgPriceList","UpdGradSlsPrc"], None),
        (["PPS040"], ["MITVEN"], "PPS040MI", "AddItemSupplier", None, None),
        (["MMS310"], ["MITLOC"], "MMS310MI", "Update", None, None),
        (["MNS150"], ["CMNUSR"], "MNS150MI", "Add", None, None),
        (["PDS010"], ["MPDWCT"], "PDS010MI", "AddWorkCenter",
         ["ChgShiftInfo","ChgShiftPatAdj"], None),
        (["CRS025"], ["CSYTAB"], "CRS025MI", None, None, "ITGR"),
        (["CRS100"], ["CSYTAB"], "CRS100MI", None, None, "SMCD"),
        (["CRS102"], ["CSYTAB"], "CRS102MI", None, None, "SDEP"),
        (["CRS633"], ["CSYTAB"], "CRS633MI", None, None, "AICL")
    ]

    # --------------------------------------------------------
    # Ensure all tables exist in m3TableCols
    # --------------------------------------------------------
    all_tables = {t for _, tables, *_ in sql_blocks for t in tables}

    for table in sorted(all_tables):
        if not table_exists_in_m3TableCols(conn, table):
            extract_table_if_missing(table)

            # 🔄 Refresh connection so SQLite sees new data
            conn.close()
            conn = sqlite3.connect(DB_PATH)

    # --------------------------------------------------------
    # Execute SQL blocks
    # --------------------------------------------------------
    for params in tqdm(sql_blocks, desc="Building m3Api2Table"):
        execute_sql_block(conn, *params)

    conn.commit()
    conn.close()
    print("✅ m3Api2Table successfully built.")

# ============================================================
if __name__ == "__main__":
    build_m3_table2api()
