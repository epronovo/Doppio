#!/usr/bin/env python3
"""
m3_repacker_db.py — Rebuild M3 Grid Access binary table exports from a SQLite database.

Companion to m3_repacker.py: instead of pairing an edited CSV with the
original binary export (to steal its schema), this reads both the record data
and the field-def schema captured by m3_unpacker_db.py (in the
m3_unpacker_header / m3_unpacker_schema tracking tables) directly from the
database, so no original binary files are needed for a round trip.
"""

import os
import sys
import sqlite3
import argparse
import zipfile

from m3_repacker import encode_bitmap_record, _is_short_decimal
from m3_unpacker import write_table_info

SCHEMA_TABLE = 'm3_unpacker_schema'
HEADER_TABLE = 'm3_unpacker_header'
DEFAULT_DB   = os.path.expanduser('~/sqlite/doppio.db')


def quote_ident(name):
    return '"' + name.replace('"', '""') + '"'


def list_tracked_tables(conn):
    rows = conn.execute(f'SELECT table_name FROM {HEADER_TABLE} ORDER BY table_name').fetchall()
    return [r[0] for r in rows]


def _has_scale_column(conn):
    return 'field_scale' in {
        r[1] for r in conn.execute(f'PRAGMA table_info({SCHEMA_TABLE})')
    }


def load_schema(conn, table_name):
    """Returns (header_bytes, fields, types, widths, scales) or None if untracked."""
    header_row = conn.execute(
        f'SELECT header FROM {HEADER_TABLE} WHERE table_name = ?', (table_name,)
    ).fetchone()
    if not header_row:
        return None

    has_scale = _has_scale_column(conn)
    scale_col = 'field_scale' if has_scale else 'NULL AS field_scale'
    field_rows = conn.execute(
        f'SELECT field_name, field_type, field_width, {scale_col} FROM {SCHEMA_TABLE} '
        'WHERE table_name = ? ORDER BY field_seq',
        (table_name,),
    ).fetchall()
    if not field_rows:
        return None

    fields = [r[0] for r in field_rows]
    types  = [r[1] for r in field_rows]
    widths = [r[2] for r in field_rows]
    scales = [r[3] for r in field_rows]
    return header_row[0], fields, types, widths, scales


def repack_table(conn, table_name, output_dir, verbose=True):
    schema = load_schema(conn, table_name)
    if schema is None:
        if verbose: print(f"  ERROR {table_name} — no schema tracked in database")
        return None

    header_bytes, fields, types, widths, scales = schema
    n_fields = len(fields)
    if any(s is None for s in scales):
        if verbose:
            print(f"  WARN  {table_name} — no field_scale tracked (unpacked before "
                  "the decimals-aware fix); re-run m3_unpacker_db.py so decimal "
                  "fields repack correctly.")
    short_decimal = [
        _is_short_decimal(types[i] or '', widths[i] or 0, scales[i] or 0)
        for i in range(n_fields)
    ]

    sql_table = table_name.lower()
    try:
        rows = conn.execute(f'SELECT * FROM {quote_ident(sql_table)} ORDER BY rowid').fetchall()
    except sqlite3.OperationalError as exc:
        if verbose: print(f"  ERROR {table_name} — data table '{sql_table}' missing: {exc}")
        return None

    payload = bytearray(header_bytes)
    for row in rows:
        # None (not stored) and '' (stored with M3's inherit marker) are
        # distinct wire states — pass through unchanged, see encode_bitmap_record.
        payload += encode_bitmap_record(list(row), n_fields, short_decimal=short_decimal)

    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, table_name)
    with open(output_path, 'wb') as f:
        f.write(bytes(payload))

    if verbose:
        print(f"  OK    {table_name}  →  {output_path}  ({len(rows)} record(s), {n_fields} field(s))")
    return output_path, len(rows)


def main():
    parser = argparse.ArgumentParser(
        description="Rebuild M3 Grid Access binary table exports from a SQLite database.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "python3  m3_repacker_db.py --db ~/sqlite/doppio.db -t CMITRN -o ./repacked/\n"
            "python3  m3_repacker_db.py --db ~/sqlite/doppio.db -o ./repacked/            (all tracked tables)\n"
            "python3  m3_repacker_db.py --db ~/sqlite/doppio.db -o ./repacked/ --zip\n"
            "python3  m3_repacker_db.py --db ~/sqlite/doppio.db -o ./repacked/ --zip --table-info ./TABLE_INFO\n"
        ),
    )
    parser.add_argument(
        '--db', dest='db', default=DEFAULT_DB,
        help=f"SQLite database path (default: {DEFAULT_DB})."
    )
    parser.add_argument(
        '-t', '--table', dest='table', default=None,
        help="Single M3 table name to repack (e.g. CMITRN). Case-insensitive. "
             "Default: repack every table tracked in the database."
    )
    parser.add_argument(
        '-o', '--output', dest='output', required=True,
        help="Output directory for repacked binaries."
    )
    parser.add_argument(
        '-q', '--quiet', action='store_true',
        help="Suppress per-file progress messages."
    )
    parser.add_argument(
        '-z', '--zip', dest='zip_path', nargs='?', const='',
        metavar='ZIP_PATH',
        help="Pack all output binaries into a zip archive. "
             "Optionally supply a path/name for the zip file; "
             "defaults to <output_dir>.zip."
    )
    parser.add_argument(
        '--table-info', dest='table_info', default=None, metavar='PATH',
        help="Path to a TABLE_INFO file to include in the zip instead of "
             "auto-generating one. By default, a TABLE_INFO listing each "
             "repacked table's actual row count is built and zipped "
             "automatically (M3 needs one for uploads); pass this to merge "
             "in counts for tables outside this batch instead."
    )

    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(1)

    args = parser.parse_args()
    verbose = not args.quiet

    if not os.path.isfile(args.db):
        print(f"ERROR: database not found: {args.db}", file=sys.stderr)
        sys.exit(2)

    if args.table_info and not os.path.isfile(args.table_info):
        print(f"ERROR: --table-info file not found: {args.table_info}", file=sys.stderr)
        sys.exit(2)

    conn = sqlite3.connect(args.db)

    try:
        conn.execute(f'SELECT 1 FROM {HEADER_TABLE} LIMIT 1')
    except sqlite3.OperationalError:
        print(f"ERROR: '{HEADER_TABLE}' not found in {args.db} — run m3_unpacker_db.py first.",
              file=sys.stderr)
        sys.exit(2)

    table_names = [args.table.upper()] if args.table else list_tracked_tables(conn)
    if not table_names:
        print("No tables tracked in database.", file=sys.stderr)
        sys.exit(2)

    if verbose:
        print(f"Batch: {len(table_names)} table(s) from '{args.db}'  →  '{args.output}'\n")

    ok = err = 0
    packed_files = []
    record_counts = []
    for table_name in table_names:
        result = repack_table(conn, table_name, args.output, verbose=verbose)
        if result:
            output_path, n_rows = result
            ok += 1
            packed_files.append(output_path)
            record_counts.append((table_name, n_rows))
        else:
            err += 1

    if verbose:
        print(f"\nDone: {ok} succeeded, {err} skipped/errored.")

    if args.zip_path is not None and packed_files:
        zip_path = args.zip_path if args.zip_path else os.path.abspath(args.output).rstrip(os.sep) + '.zip'
        # Files land at the root of the archive (no nested directory) — M3
        # rejects uploads where files are nested inside a directory.
        with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
            for fpath in packed_files:
                zf.write(fpath, arcname=os.path.basename(fpath))
            if args.table_info:
                zf.write(args.table_info, arcname='TABLE_INFO')
                ti_note = "TABLE_INFO (supplied)"
            else:
                zf.writestr('TABLE_INFO', write_table_info(record_counts))
                ti_note = f"TABLE_INFO (auto: {', '.join(f'{t}={n}' for t, n in record_counts)})"
        if verbose:
            print(f"ZIP   {zip_path}  ({len(packed_files)} file(s) + {ti_note})")

    conn.close()


if __name__ == '__main__':
    main()
