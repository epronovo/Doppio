# MpGd_Build_m3Api2Table.py
# -----------------------------------------------------------------------
# PURPOSE
#   Builds the m3Api2Table by joining M3 API field definitions (cmifld)
#   with table/column metadata (m3TableCols) and screen field positions
#   (CSEFPU), driven by rules in the ExtractRules table.  Each output row
#   maps an M3 API transaction field to its corresponding database table
#   column, UI program, and panel.  Existing rows are never duplicated;
#   the INSERT is guarded by a NOT EXISTS check on the sequence.
#
# INPUTS
#   - ExtractRules  (SQLite) — defines which APIs, programs, and tables
#                              to process, plus optional ignore/blank rules
#   - cmifld        (SQLite) — M3 API field definitions (API, transaction,
#                              direction, field name, position)
#   - m3TableCols   (SQLite) — M3 table column metadata (populated by
#                              MpGd_Extract_M3Info)
#   - CSEFPU        (SQLite) — screen field-to-program position mappings
#
# OUTPUTS
#   - m3Api2Table   (SQLite) — API ↔ table/column mapping
#
# DEPENDENCIES
#   - config.py
# -----------------------------------------------------------------------

import logging
import sqlite3
from tqdm import tqdm

from config import get_sqlite_db_path

logger = logging.getLogger(__name__)

DB_PATH = get_sqlite_db_path()

# ============================================================
# CREATE TARGET TABLE IF NEEDED
# ============================================================
CREATE_M3API2TABLE_SQL = """
CREATE TABLE IF NOT EXISTS m3Api2Table (
    API             TEXT,
    TransactionName TEXT,
    Direction   	TEXT,
    FieldName   	TEXT,
    TableName   	TEXT,
    ColumnName  	TEXT,
    Program     	TEXT,
    Panel       	TEXT,
    FRPO        	INTEGER,
    Sequence    	INTEGER
);
"""

# ============================================================
# INSERT SQL (NAMED PARAMETERS)
# ============================================================
INSERT_SQL = """
INSERT INTO m3Api2Table
SELECT * FROM (
	SELECT
	    a.MINM                  AS API,
	    MAX(a.TRNM)             AS TransactionName,
	    a.TRTP                  AS Direction,
	    a.FLNM                  AS FieldName,
	    :table                      AS TableName,
	    COALESCE(t.ColumnName,'')   AS ColumnName,
	    COALESCE(t.j3pgm,'')        AS j3pgm,
	    COALESCE(t.j3pic1,'')       AS j3pic1,
	    CAST(a.FRPO AS INTEGER)     AS FRPO,
	    COALESCE(t.sequence,0)      AS Sequence
	FROM cmifld a
	LEFT JOIN (
		SELECT
		    mg.TableName                              AS TableName,
		    mg.ColumnName                             AS ColumnName,
		    mg.Description                            AS Description,
		    mg.DataType                               AS DataType,
		    mg.Length                                 AS Length,
		    COALESCE(mg.Decimals,'')                  AS Decimals,
		    COALESCE(c.j3pgm,'')                      AS j3pgm,
		    COALESCE(MIN(c.j3pic1),'')                AS j3pic1,
		    CASE 
		        WHEN mg.TableName = 'OCUSMA' THEN 
		            REPLACE(SUBSTR(COALESCE(mg.ColumnName,''), -4), 'LHCD', 'LNCD')
		        WHEN mg.TableName = 'OCUSAD' THEN 
		            REPLACE(SUBSTR(COALESCE(mg.ColumnName,''), -4), 'EDES', 'EDE2')
		        ELSE 
		            SUBSTR(COALESCE(mg.ColumnName,''), -4)
		    END                                       AS FieldName,
		    mg.Sequence                               AS Sequence
		FROM m3TableCols mg
		LEFT JOIN CSEFPU c
		    ON SUBSTR(COALESCE(c.j3fldi,''),-4)
		     = SUBSTR(COALESCE(mg.ColumnName,''),-4)
		   AND c.j3pgm = :pgm
		   AND c.j3pic1 NOT IN ('A','B','C','D','P')
		WHERE mg.TableName = substr( :table, 1, 6 )
		GROUP BY mg.TableName,mg.ColumnName,mg.Description,mg.DataType,mg.Length,mg.Decimals,mg.Sequence,c.j3pgm
	) t ON t.FieldName = a.FLNM
	LEFT JOIN m3Api2Table b on b.API = a.MINM AND b.TransactionName = a.TRNM AND b.Direction = a.TRTP AND b.FieldName = a.FLNM 
	LEFT JOIN m3Api2Table c on c.sequence = t.sequence
	WHERE a.MINM = :api AND a.TRNM LIKE :transaction AND a.TRNM NOT LIKE :ignore 
	AND b.sequence is null AND c.sequence is null
	GROUP BY a.MINM,a.TRTP,a.FLNM,t.TableName,t.ColumnName,t.sequence
) final
WHERE ( :IncludeBlanks = 'yes' OR TableName <> '' )
ORDER BY API,TransactionName,Direction,FRPO;
"""

# ============================================================
# MAIN DRIVER
# ============================================================
def build_m3Api2Table():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # Ensure table exists
    cur.execute(CREATE_M3API2TABLE_SQL)

    # Load rules
    rules = cur.execute("""
        SELECT
            Pgm,
            Tbl,
            API,
            API_Transaction,
            COALESCE(IgnoreRule,'')      AS IgnoreRule,
            COALESCE(IncludeBlanks,'no') AS IncludeBlanks
        FROM ExtractRules
        ORDER BY Seq
    """).fetchall()

    # --------------------------------------------------------
    # Calculate total steps for progress bar
    # --------------------------------------------------------
    total_steps = 0
    expanded_rules = []

    for r in rules:
        pgms = [p.strip() for p in r["Pgm"].split(",")]
        tables = [t.strip() for t in r["Tbl"].split(",")]
        for pgm in pgms:
            for table in tables:
                expanded_rules.append((r, pgm, table))
                total_steps += 1

    total_inserted = 0

    # --------------------------------------------------------
    # Execute with progress bar
    # --------------------------------------------------------
    with tqdm(total=total_steps, desc="Building m3Api2Table", unit="step") as pbar:
        for r, pgm, table in expanded_rules:
            params = {
                "pgm": pgm,
                "table": table,
                "api": r["API"],
                "transaction": r["API_Transaction"],
                "ignore": r["IgnoreRule"],
                "IncludeBlanks": r["IncludeBlanks"]
            }

            cur.execute(INSERT_SQL, params)
            total_inserted += cur.rowcount
            pbar.update(1)

    conn.commit()
    conn.close()

    logger.info(f"✅ m3Api2Table populated — {total_inserted} rows inserted.")


# ============================================================
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S"
    )
    build_m3Api2Table()
