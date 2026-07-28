#!/usr/bin/env python3
"""
m3_repacker.py — Convert edited CSV files back to M3 Grid Access binary tables.

Round-trip companion to m3_unpacker.py. Uses a "Read-Modify-Write" approach:
it steals the exact binary schema (field definitions) from the original M3
export and stitches it onto your modified CSV data, then emits records in the
same bitmap-based layout M3 uses natively:

    Per record:
      [4 bytes]  record block size (big-endian u32) = bitmap + all LPS data
      [K bytes]  nullability bitmap, K = ceil(n_fields / 8), MSB-first.
                 Bit i = 1 → field i is present; 0 → null/empty (omitted).
      [...]      For each present field: 4-byte big-endian length + UTF-8 bytes.

A CSV cell that is empty ('') is treated as null (bit cleared, value omitted),
matching how m3_unpacker.py decodes empty cells.
"""

import struct
import csv
import sys
import os
import argparse
import zipfile


# CSV token for a field that is NOT stored on a record (bitmap bit clear), as
# emitted by m3_unpacker.py in raw mode. Distinct from a plain empty cell, which
# means the field IS stored but empty (LPS length-0 "inherit"). Keep in sync
# with m3_unpacker.NULL_SENTINEL.
NULL_SENTINEL = r'\N'


# ─────────────────────────────────────────────
# Field definition parser (mirrors m3_unpacker.py)
# ─────────────────────────────────────────────

def parse_field_defs(data):
    """Returns (fields, types, widths, scales, header_length). header_length = 4 + fdef_length.
    scales are the per-field decimal places (4th field-def component)."""
    if len(data) < 4:
        raise ValueError("File too short.")
    fdef_length = struct.unpack('>I', data[:4])[0]
    if 4 + fdef_length > len(data):
        raise ValueError(f"Field def length {fdef_length} exceeds file size.")

    fdef_bytes = data[4:4 + fdef_length]
    fields = []
    types  = []
    widths = []
    scales = []
    for entry in fdef_bytes.split(b'\x01'):
        entry = entry.strip(b'\x00')
        if not entry:
            continue
        try:
            parts = entry.decode('ascii').split(';')
            if len(parts) >= 3:
                types.append(parts[0].strip())
                fields.append(parts[1].strip())
                try:
                    widths.append(int(parts[2].strip()))
                except ValueError:
                    widths.append(0)
                scale = 0
                if len(parts) >= 4:
                    try:
                        scale = int(parts[3].strip())
                    except ValueError:
                        scale = 0
                scales.append(scale)
        except (UnicodeDecodeError, ValueError):
            continue

    return fields, types, widths, scales, 4 + fdef_length


# JDBC numeric type codes and the short-decimal cutoff, mirrored from
# m3_unpacker.py: whole-number fields (scale 0) of these types with display
# width <= 9 are stored as a raw 4-byte big-endian int (no LPS length prefix).
# Fields with decimal places (scale > 0) are always LPS strings.
_NUMERIC_TYPES = {'4', '-5', '5', '-6', '7', '6', '8', '2', '3'}
_SHORT_DECIMAL_MAX_WIDTH = 9


def _is_short_decimal(type_code, width, scale=0):
    return (
        type_code in _NUMERIC_TYPES
        and (scale or 0) == 0
        and width <= _SHORT_DECIMAL_MAX_WIDTH
    )


# ─────────────────────────────────────────────
# Record encoder (bitmap-style, matches m3_unpacker.py)
# ─────────────────────────────────────────────

def encode_bitmap_record(row, n_fields, short_decimal=None):
    """
    Encode one row as a bitmap-style record.

    A None cell means the field wasn't stored on the original record (bit=0,
    value omitted). An empty-string cell ('') means the field WAS stored but
    with M3's "inherit" marker (bit=1, LPS length=0) — only meaningful for
    LPS-encoded fields; for short_decimal fields (no LPS concept) '' is
    treated the same as None, since there's no wire representation for an
    empty-but-present integer. This mirrors m3_unpacker.py's
    parse_records_bitmap_style(fill=False) exactly, so a raw-unpacked table
    round-trips back to the original bytes.

    short_decimal: optional list of n_fields bools. When short_decimal[i] is
    True, field i is written as a raw 4-byte big-endian int (no LPS length
    prefix), matching how m3_unpacker.py decodes short decimal fields.
    """
    # Normalize row length to exactly n_fields
    if len(row) < n_fields:
        row = list(row) + [''] * (n_fields - len(row))
    else:
        row = list(row[:n_fields])

    if short_decimal is None:
        short_decimal = [False] * n_fields

    bitmap_len = (n_fields + 7) // 8
    bitmap = bytearray(bitmap_len)
    data_bytes = bytearray()

    for i, val in enumerate(row):
        if val is None or (val == '' and short_decimal[i]):
            # Not stored on the original record: leave the bit clear and
            # omit the value. ('' has no wire form for a short_decimal field,
            # so it's treated the same as None.)
            continue
        # Present: set the bit (MSB-first within each byte) and write the value.
        byte_idx = i // 8
        bit_idx  = 7 - (i % 8)
        bitmap[byte_idx] |= (1 << bit_idx)

        if short_decimal[i]:
            # Raw 4-byte big-endian signed int — no LPS prefix.
            try:
                int_val = int(str(val).strip())
            except ValueError:
                int_val = int(float(str(val).strip()))
            data_bytes += struct.pack('>i', int_val)
            continue

        if val == '':
            # M3's "inherit" marker: bit set, LPS length=0, no value bytes.
            data_bytes += struct.pack('>I', 0)
            continue

        b_val = str(val).encode('utf-8')
        data_bytes += struct.pack('>I', len(b_val)) + b_val

    body = bytes(bitmap) + bytes(data_bytes)
    return struct.pack('>I', len(body)) + body


# ─────────────────────────────────────────────
# Header extraction (original binary's schema)
# ─────────────────────────────────────────────

def extract_schema(original_bin_path):
    """Returns (header_bytes, fields_list, types_list, widths_list, scales_list)
    extracted from the original binary."""
    with open(original_bin_path, 'rb') as f:
        data = f.read()

    fields, types, widths, scales, header_len = parse_field_defs(data)
    if not fields:
        raise ValueError(f"No field definitions found in {original_bin_path}.")
    return data[:header_len], fields, types, widths, scales


# ─────────────────────────────────────────────
# Core repacker
# ─────────────────────────────────────────────

def repack_file(input_csv, original_bin, output_bin=None, verbose=True):
    if output_bin is None:
        # Same base name as the CSV, no extension.
        output_bin = os.path.splitext(input_csv)[0]

    if not os.path.isfile(input_csv):
        if verbose: print(f"  ERROR {input_csv} — CSV not found")
        return False

    if not os.path.isfile(original_bin):
        if verbose: print(f"  ERROR {original_bin} — original binary not found")
        return False

    # 1. Steal the original schema (field defs header + field names/types/widths/scales)
    try:
        header_bytes, fields, types, widths, scales = extract_schema(original_bin)
    except Exception as e:
        if verbose: print(f"  ERROR {original_bin} — {e}")
        return False

    n_fields = len(fields)
    short_decimal = [
        _is_short_decimal(
            types[i] if i < len(types) else '',
            widths[i] if i < len(widths) else 0,
            scales[i] if i < len(scales) else 0,
        )
        for i in range(n_fields)
    ]

    # 2. Read CSV (pipe-delimited, matching m3_unpacker.py output). The NULL
    # sentinel marks a field that was absent on the original record (bitmap bit
    # clear); map it back to None so encode_bitmap_record omits it. A plain
    # empty cell stays '' — a present-but-empty field.
    with open(input_csv, 'r', newline='', encoding='utf-8') as f:
        reader = csv.reader(f, delimiter='|')
        try:
            csv_headers = next(reader)
        except StopIteration:
            if verbose: print(f"  SKIP  {input_csv} — CSV is empty")
            return False

        rows = [
            [None if c == NULL_SENTINEL else c for c in r]
            for r in reader
            if any(c != '' for c in r)
        ]

    # Optional sanity check: warn if CSV header doesn't match binary schema
    if verbose and [h.strip() for h in csv_headers[:n_fields]] != fields:
        mismatch_count = sum(
            1 for a, b in zip([h.strip() for h in csv_headers], fields) if a != b
        )
        if mismatch_count:
            print(f"  WARN  {os.path.basename(input_csv)} — CSV header differs "
                  f"from binary schema in {mismatch_count} column(s); using binary schema order")

    # 3. Reconstruct payload (original schema header + new records)
    payload = bytearray(header_bytes)
    for row in rows:
        payload += encode_bitmap_record(row, n_fields, short_decimal=short_decimal)

    # 4. Write output
    os.makedirs(os.path.dirname(os.path.abspath(output_bin)), exist_ok=True)
    with open(output_bin, 'wb') as f:
        f.write(payload)

    if verbose:
        print(f"  OK    {os.path.basename(input_csv)}  →  {os.path.basename(output_bin)} "
              f"({len(rows)} record(s), {n_fields} field(s))")
    return True


# ─────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Repack pipe-delimited CSVs back into M3 Grid Access binary tables.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "python3  m3_repacker.py -i edited.csv -b original/CMNDIV -o CMNDIV\n"
            "python3  m3_repacker.py -i ./csv_output/ -b ./cmp001_tables/ -o ./csv_output/repacked/\n"
            "python3  m3_repacker.py -i ./csv_output/ -b ./cmp001_tables/ -o ./repacked/ --zip\n"
            "python3  m3_repacker.py -i ./csv_output/ -b ./cmp001_tables/ --zip ./my_export.zip\n"
            "\n"
            "Output file names mirror the input (same base name, no extension).\n"
            "Batch mode pairs files by base name: <csv_dir>/FOO.csv ↔ <bin_dir>/FOO\n"
            "--zip / -z packs all repacked binaries into a zip archive.\n"
            "  With no value: auto-names the zip (<output_dir>.zip or <output_file>.zip).\n"
            "  With a value:  writes to the specified path.\n"
        ),
    )
    parser.add_argument(
        '-i', '--input', dest='input', required=True,
        help="Input CSV file OR directory of CSVs."
    )
    parser.add_argument(
        '-b', '--binary', dest='binary', required=True,
        help="Original M3 binary file (schema source) OR directory of originals."
    )
    parser.add_argument(
        '-o', '--output', dest='output', default=None,
        help="Output file (file mode) or directory (batch mode). "
             "Output file names mirror the input CSV base name with no extension. "
             "Defaults to the CSV's base name (file mode) or <csv_dir>/repacked/ (batch)."
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
             "defaults to <output_dir>.zip (batch) or <output_file>.zip (single)."
    )

    # When run with no arguments, show the full help (with examples) instead
    # of argparse's terse "error: the following arguments are required" line.
    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(1)

    args = parser.parse_args()
    verbose = not args.quiet

    csv_arg  = args.input
    bin_arg  = args.binary
    out_arg  = args.output
    zip_arg  = args.zip_path   # None → no zip; '' → auto-name; str → explicit path

    if not os.path.exists(csv_arg):
        print(f"ERROR: input path does not exist: {csv_arg}", file=sys.stderr)
        sys.exit(2)
    if not os.path.exists(bin_arg):
        print(f"ERROR: binary path does not exist: {bin_arg}", file=sys.stderr)
        sys.exit(2)

    # Batch mode
    if os.path.isdir(csv_arg) and os.path.isdir(bin_arg):
        out_dir = out_arg or os.path.join(csv_arg, "repacked")
        os.makedirs(out_dir, exist_ok=True)

        csv_files = sorted(f for f in os.listdir(csv_arg) if f.endswith('.csv'))

        if verbose:
            print(f"Batch: {len(csv_files)} file(s) in '{csv_arg}'  →  '{out_dir}'\n")

        ok = err = 0
        packed_files = []
        for csv_name in csv_files:
            base_name = os.path.splitext(csv_name)[0]
            csv_path  = os.path.join(csv_arg, csv_name)
            bin_path  = os.path.join(bin_arg, base_name)
            out_path  = os.path.join(out_dir, base_name)

            if repack_file(csv_path, bin_path, out_path, verbose=verbose):
                ok += 1
                packed_files.append(out_path)
            else:
                err += 1

        if verbose:
            print(f"\nDone: {ok} succeeded, {err} skipped/errored.")

        if zip_arg is not None and packed_files:
            zip_path = zip_arg if zip_arg else os.path.abspath(out_dir).rstrip(os.sep) + '.zip'
            # Files are written at the root of the archive (arcname is just
            # the basename). M3 rejects uploads where files are nested inside
            # a directory.
            with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                for fpath in packed_files:
                    zf.write(fpath, arcname=os.path.basename(fpath))
                # Always include TABLE_INFO from the original binary directory
                # if present — M3 needs it alongside the data tables.
                ti_src = os.path.join(bin_arg, 'TABLE_INFO')
                if os.path.isfile(ti_src):
                    zf.write(ti_src, arcname='TABLE_INFO')
            if verbose:
                extra = " (+ TABLE_INFO)" if os.path.isfile(os.path.join(bin_arg, 'TABLE_INFO')) else ""
                print(f"ZIP   {zip_path}  ({len(packed_files)} file(s){extra})")

    # Single-file mode
    else:
        if os.path.isdir(csv_arg) or os.path.isdir(bin_arg):
            print("ERROR: --input and --binary must both be files or both be directories.",
                  file=sys.stderr)
            sys.exit(2)
        success = repack_file(csv_arg, bin_arg, out_arg, verbose=verbose)

        if zip_arg is not None and success:
            # Resolve the actual output path the same way repack_file() does
            resolved_out = out_arg if out_arg else os.path.splitext(csv_arg)[0]
            zip_path = zip_arg if zip_arg else resolved_out + '.zip'
            # arcname uses just the basename so files land at the root of the
            # archive (M3 rejects archives with a nested directory).
            with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                zf.write(resolved_out, arcname=os.path.basename(resolved_out))
                # Pull TABLE_INFO from the same directory as the original binary.
                ti_src = os.path.join(os.path.dirname(os.path.abspath(bin_arg)), 'TABLE_INFO')
                if os.path.isfile(ti_src):
                    zf.write(ti_src, arcname='TABLE_INFO')
            if verbose:
                ti_src = os.path.join(os.path.dirname(os.path.abspath(bin_arg)), 'TABLE_INFO')
                extra = " (+ TABLE_INFO)" if os.path.isfile(ti_src) else ""
                print(f"ZIP   {zip_path}{extra}")


if __name__ == '__main__':
    main()
