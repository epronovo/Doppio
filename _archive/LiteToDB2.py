# main.py
import os
import sqlite3
import pyodbc

SQLITE_DB = "C:/sqlite/asr.db"
DB2_CONN_STR = (
    "DSN=Raymond DB;"
    "UID=ericp;"
    "PWD=EPM3con531;"
)
QUERIES_DIR = "toDB2"


def read_sql_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        return f.read()


def sqlite_query_to_df(sql):
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    cursor.execute(sql)
    columns = [desc[0] for desc in cursor.description]
    rows = cursor.fetchall()
    conn.close()
    return columns, rows


def map_sqlite_to_db2_type(sqlite_type):
    if "INT" in sqlite_type.upper():
        return "INTEGER"
    elif "CHAR" in sqlite_type.upper() or "TEXT" in sqlite_type.upper():
        return "VARCHAR(255)"
    elif "REAL" in sqlite_type.upper() or "FLOAT" in sqlite_type.upper():
        return "FLOAT"
    else:
        return "VARCHAR(255)"


def infer_column_types_from_sqlite(sql):
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    cursor.execute(f"SELECT * FROM ({sql}) LIMIT 1")
    col_info = cursor.description
    types = []
    for col in col_info:
        types.append("VARCHAR(255)")  # Conservative default
    conn.close()
    return types


def create_db2_table(cursor, table_name, columns, types):
    # cursor.execute(f"DROP TABLE qgpl.{table_name}")
    column_defs = ", ".join(
        f"{col} {col_type}" for col, col_type in zip(columns, types)
    )
    # print(f"CREATE TABLE qgpl.{table_name} ({column_defs})")
    # cursor.execute(f"CREATE TABLE qgpl.{table_name} ({column_defs})")
    # cursor.execute("CREATE TABLE qgpl.CAM (CUSN05 VARCHAR(255), DSEQ05 VARCHAR(255), CUNO VARCHAR(255), ADID VARCHAR(255))")
    # cursor.execute("INSERT INTO qgpl.CAM (CUSN05, DSEQ05, CUNO, ADID) VALUES ('10283', '10283', '10283', '10283')")

def insert_into_db2(cursor, table_name, columns, rows):
    placeholders = ", ".join(["?"] * len(columns))
    column_names = ", ".join(columns)
    sql = f"INSERT INTO qgpl.{table_name} ({column_names}) VALUES ({placeholders})"

    output_file = f"{table_name}_inserts.txt"
    with open(output_file, "w", encoding="utf-8") as f:
        for row in rows:
            values_str = ", ".join(
                [f"'{str(v).replace('\'', '\'\'')}'" if v is not None else "NULL" for v in row]
            )
            sql_server_insert = f"INSERT INTO qgpl.{table_name} ({column_names}) VALUES ({values_str});"
            f.write(sql_server_insert + "\n")
            # cursor.execute(sql, row)


def main():
    db2_conn = pyodbc.connect(DB2_CONN_STR)
    db2_cursor = db2_conn.cursor()

    for filename in os.listdir(QUERIES_DIR):
        if filename.endswith(".sql"):
            path = os.path.join(QUERIES_DIR, filename)
            table_name = os.path.splitext(filename)[0]

            sql = read_sql_file(path)
            columns, rows = sqlite_query_to_df(sql)
            types = infer_column_types_from_sqlite(sql)

            print(f"Transferring data to DB2 table qgpl.{table_name}...")

            try:
                create_db2_table(db2_cursor, table_name, columns, types)
            except Exception as e:
                print(f"Failed to create table: {e}")

            try:
                insert_into_db2(db2_cursor, table_name, columns, rows)
                db2_conn.commit()
            except Exception as e:
                print(f"Insert failed: {e}")
                db2_conn.rollback()

    db2_conn.close()


if __name__ == "__main__":
    main()