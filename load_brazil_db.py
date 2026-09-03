#!/usr/bin/env python3
r"""
load_brazil_db.py — Replace the row data of the tables tracked by
m3_unpacker_db.py with the contents of new pipe-delimited CSVs.

  python3 load_brazil_db.py <db_path> <csv_dir>

The schema/header tracking rows (m3_unpacker_schema, m3_unpacker_header) are
left untouched, so m3_repacker_db.py still rebuilds a byte-accurate M3 export
header — only the records change. The CSV's \N sentinel maps back to NULL
(bitmap bit clear), exactly as m3_repacker.py reads it.
"""

import csv
import os
import sqlite3
import sys

NULL_SENTINEL = r'\N'
SCHEMA_TABLE = 'm3_unpacker_schema'


def quote_ident(name):
    return '"' + name.replace('"', '""') + '"'


def load(conn, csv_path, table_name):
    fields = [
        r[0] for r in conn.execute(
            f'SELECT field_name FROM {SCHEMA_TABLE} WHERE table_name = ? ORDER BY field_seq',
            (table_name,),
        )
    ]
    if not fields:
        sys.exit(f'ERROR: {table_name} is not tracked in the database.')

    with open(csv_path, newline='', encoding='utf-8') as f:
        reader = csv.reader(f, delimiter='|')
        header = [h.strip() for h in next(reader)]
        if header != fields:
            sys.exit(f'ERROR: {csv_path} header does not match the {table_name} '
                     f'binary schema.\n  csv:    {header}\n  binary: {fields}')
        rows = [[None if c == NULL_SENTINEL else c for c in r] for r in reader]

    bad = [i for i, r in enumerate(rows, 2) if len(r) != len(fields)]
    if bad:
        sys.exit(f'ERROR: {csv_path} lines {bad[:5]} have the wrong column count.')

    sql_table = table_name.lower()
    conn.execute(f'DELETE FROM {quote_ident(sql_table)}')
    conn.executemany(
        f'INSERT INTO {quote_ident(sql_table)} VALUES ({", ".join("?" * len(fields))})',
        rows,
    )
    print(f'  {table_name}: replaced with {len(rows)} row(s) from {os.path.basename(csv_path)}')
    return len(rows)


def main():
    db, csv_dir = sys.argv[1], sys.argv[2]
    conn = sqlite3.connect(db)
    print(f'Loading into {db}')
    for table_name in sorted(
        r[0] for r in conn.execute(f'SELECT DISTINCT table_name FROM {SCHEMA_TABLE}')
    ):
        path = os.path.join(csv_dir, table_name + '.csv')
        if os.path.isfile(path):
            load(conn, path, table_name)
        else:
            print(f'  {table_name}: no CSV found, leaving as-is')
    conn.commit()
    conn.close()


if __name__ == '__main__':
    main()
