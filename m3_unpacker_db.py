#!/usr/bin/env python3
"""
m3_unpacker_db.py — Parse M3 Grid Access binary table export files into a SQLite database.

Companion to m3_unpacker.py: same binary parsing (field defs, bitmap records,
0xFF 0xE0 records), but the destination is a SQLite database instead of CSV
files. Each M3 table becomes a SQLite table (lower-cased name, TEXT columns
named after the M3 field names) — matching the existing convention already
used in doppio.db (e.g. CMITRN → cmitrn).

The exact field-def header bytes and per-field type/width metadata are stored
in two tracking tables (m3_unpacker_header, m3_unpacker_schema) so that
m3_repacker_db.py can rebuild a byte-accurate binary export straight from the
database, without needing the original export files around.
"""

import os
import sys
import sqlite3
import argparse

from m3_unpacker import (
    parse_table_info,
    parse_field_defs,
    detect_format,
    parse_records_bitmap_style,
    parse_records_ffe0_style,
)

DEFAULT_DB = os.path.expanduser('~/sqlite/doppio.db')

SCHEMA_TABLE = 'm3_unpacker_schema'
HEADER_TABLE = 'm3_unpacker_header'


def quote_ident(name):
    return '"' + name.replace('"', '""') + '"'


def ensure_meta_tables(conn):
    conn.execute(f'''
        CREATE TABLE IF NOT EXISTS {SCHEMA_TABLE} (
            table_name  TEXT NOT NULL,
            field_seq   INTEGER NOT NULL,
            field_name  TEXT NOT NULL,
            field_type  TEXT,
            field_width INTEGER,
            field_scale INTEGER,
            PRIMARY KEY (table_name, field_seq)
        )
    ''')
    conn.execute(f'''
        CREATE TABLE IF NOT EXISTS {HEADER_TABLE} (
            table_name TEXT PRIMARY KEY,
            header     BLOB NOT NULL
        )
    ''')
    # Migrate older databases that predate the field_scale column. Without it,
    # decimal fields (scale > 0) can't be told apart from raw 4-byte ints on
    # repack, so any table tracked before this column existed must be re-unpacked.
    cols = {r[1] for r in conn.execute(f'PRAGMA table_info({SCHEMA_TABLE})')}
    if 'field_scale' not in cols:
        conn.execute(f'ALTER TABLE {SCHEMA_TABLE} ADD COLUMN field_scale INTEGER')


def load_table(conn, input_path, table_info=None, verbose=True, fill=True):
    """Parse one M3 binary export file and load it into the database.

    Mirrors m3_unpacker.parse_file(), but writes to a SQLite table (replacing
    it) plus the schema/header tracking tables instead of a CSV file.
    """
    with open(input_path, 'rb') as f:
        data = f.read()

    if len(data) < 8:
        if verbose: print(f"  SKIP  {input_path}  —  file too small")
        return False

    try:
        fields, types, widths, scales, data_section_start = parse_field_defs(data)
    except Exception as exc:
        if verbose: print(f"  ERROR {input_path}  —  {exc}")
        return False

    if not fields:
        if verbose: print(f"  SKIP  {input_path}  —  no field definitions")
        return False

    header_bytes = data[:data_section_start]
    n = len(fields)
    data_section = data[data_section_start:]
    table_name   = os.path.basename(input_path)
    expected     = (table_info or {}).get(table_name)
    fmt          = detect_format(data_section)

    if expected == 0:
        if verbose: print(f"  SKIP  {table_name}  —  0 records (per TABLE_INFO)")
        return False

    if fmt == 'ffe0':
        records = parse_records_ffe0_style(data_section, n)
    else:
        records = parse_records_bitmap_style(
            data_section, n, field_types=types, field_widths=widths,
            field_scales=scales, expected_count=expected, fill=fill,
        )

    if not fill:
        records = [r for r in records if any(v for v in r)]

    if fmt == 'gap' and expected and len(records) > expected:
        records = records[:expected]

    if not records:
        if verbose: print(f"  SKIP  {table_name}  —  0 records parsed")
        return False

    sql_table = table_name.lower()

    conn.execute(f'DROP TABLE IF EXISTS {quote_ident(sql_table)}')
    cols_sql = ', '.join(f'{quote_ident(fn)} TEXT' for fn in fields)
    conn.execute(f'CREATE TABLE {quote_ident(sql_table)} ({cols_sql})')

    placeholders = ', '.join('?' for _ in fields)
    conn.executemany(
        f'INSERT INTO {quote_ident(sql_table)} VALUES ({placeholders})',
        records,
    )

    conn.execute(f'DELETE FROM {SCHEMA_TABLE} WHERE table_name = ?', (table_name,))
    conn.executemany(
        f'INSERT INTO {SCHEMA_TABLE} '
        '(table_name, field_seq, field_name, field_type, field_width, field_scale) '
        'VALUES (?, ?, ?, ?, ?, ?)',
        [
            (table_name, i, fields[i],
             types[i] if i < len(types) else '',
             widths[i] if i < len(widths) else 0,
             scales[i] if i < len(scales) else 0)
            for i in range(n)
        ],
    )
    conn.execute(
        f'INSERT INTO {HEADER_TABLE} (table_name, header) VALUES (?, ?) '
        'ON CONFLICT(table_name) DO UPDATE SET header = excluded.header',
        (table_name, header_bytes),
    )

    count_note = f" (expected {expected})" if expected else ""
    if verbose:
        print(f"  OK    {table_name}  →  table '{sql_table}'  ({len(records)} record(s){count_note}, {n} field(s))")
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Parse M3 Grid Access binary table export files into a SQLite database.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "python3  m3_unpacker_db.py -i ./cmp200_tables/ --db ~/sqlite/doppio.db\n"
            "python3  m3_unpacker_db.py --input TABLE.bin --db ~/sqlite/doppio.db\n"
            "python3  m3_unpacker_db.py -i /path/to/input_dir              (uses default db)\n"
        ),
    )
    parser.add_argument(
        '-i', '--input', dest='input', required=True,
        help="Input file OR directory containing M3 binary table exports."
    )
    parser.add_argument(
        '--db', dest='db', default=DEFAULT_DB,
        help=f"SQLite database path (default: {DEFAULT_DB})."
    )
    parser.add_argument(
        '-q', '--quiet', action='store_true',
        help="Suppress per-file progress messages."
    )
    parser.add_argument(
        '--fill', action='store_true',
        help="Fill in inherited values from previous rows and type defaults "
             "(0 for numeric) for fields not stored on a record, instead of "
             "leaving them empty. Makes the table more readable, but the "
             "filled-in values are indistinguishable from real data once "
             "stored, so m3_repacker_db.py can no longer tell which fields "
             "were actually absent — it will mark them present and produce "
             "a binary export M3 rejects. Do not use this if the table will "
             "ever be repacked."
    )

    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(1)

    args = parser.parse_args()
    verbose = not args.quiet
    fill    = args.fill

    if not os.path.exists(args.input):
        print(f"ERROR: input path does not exist: {args.input}", file=sys.stderr)
        sys.exit(2)

    if os.path.isdir(args.input):
        ti_path = os.path.join(args.input, 'TABLE_INFO')
    else:
        ti_path = os.path.join(os.path.dirname(os.path.abspath(args.input)), 'TABLE_INFO')

    table_info = parse_table_info(ti_path)
    if table_info and verbose:
        print(f"TABLE_INFO loaded: {table_info}")

    os.makedirs(os.path.dirname(os.path.abspath(args.db)), exist_ok=True)
    conn = sqlite3.connect(args.db)
    ensure_meta_tables(conn)

    if os.path.isdir(args.input):
        files = sorted(
            f for f in os.listdir(args.input)
            if os.path.isfile(os.path.join(args.input, f))
            and not f.startswith('.')
            and f not in ('TABLE_INFO',)
            and not f.endswith('.csv')
        )
        if verbose:
            print(f"Batch: {len(files)} file(s) in '{args.input}'  →  '{args.db}'\n")
        ok = err = 0
        for fname in files:
            if load_table(
                conn, os.path.join(args.input, fname),
                table_info=table_info, verbose=verbose, fill=fill,
            ):
                ok += 1
            else:
                err += 1
        conn.commit()
        if verbose:
            print(f"\nDone: {ok} succeeded, {err} skipped/errored.")
    else:
        if load_table(conn, args.input, table_info=table_info, verbose=verbose, fill=fill):
            conn.commit()

    conn.close()


if __name__ == '__main__':
    main()
