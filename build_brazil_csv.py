#!/usr/bin/env python3
r"""
build_brazil_csv.py — Turn the 'ZITMAS DATA' / 'ZIDADR DATA' tabs of the Brazil
migration workbook into pipe-delimited CSVs that match the EXACT field layout of
the M3 XITMAS / XIDADR binary exports, ready for m3_repacker.py / m3_repacker_db.py.

  python3 build_brazil_csv.py <workbook.xlsx> <out_csv_dir> <orig_binary_dir>

Key rules (derived from the live MBR/XITMAS + MBR/XIDADR binaries):

  * The M3 tables have MORE fields than the source tabs (XITMAS 39 vs 25,
    XIDADR 16 vs 12). The sheet columns line up with the leading fields in
    order; every trailing M3-only field is written as \N (bitmap bit clear =
    "not stored", which M3 renders as its type default).

  * NEVER emit a bare '' cell. In M3's wire format an LPS length of 0 is the
    "inherit from the previous record" marker, not a blank. A genuinely blank
    char value is written as \N instead.

  * Char/varchar fields are nchar-padded in the source; strip TRAILING ASCII
    spaces only. Leading spaces and NBSP are preserved — in SQL Server those
    make ' 332A24-6026-20' a genuinely different row from '332A24-6026-20',
    and collapsing them would create duplicate M3 keys.

  * Numeric, scale 0, width <= 9  -> raw 4-byte int on the wire; plain integer.
  * Numeric, scale > 0           -> LPS string with that many decimals ("0.00").
  * Numeric, scale 0, width >= 10 -> LPS string; plain digits.
"""

import csv
import os
import sys

import openpyxl

from m3_unpacker import parse_field_defs, _NUMERIC_TYPES, _is_short_decimal

NULL = r'\N'

TABS = {
    'ZITMAS DATA': 'XITMAS',
    'ZIDADR DATA': 'XIDADR',
}


def norm_cell(value, type_code, width, scale):
    """Render one source cell as the CSV token m3_repacker.py expects."""
    if value is None:
        return NULL

    if type_code in _NUMERIC_TYPES:
        if isinstance(value, str):
            value = value.strip()
            if value == '':
                return NULL
        if _is_short_decimal(type_code, width, scale):
            return str(int(float(value)))
        if scale:
            return f'{float(value):.{scale}f}'
        return str(int(float(value)))

    if not isinstance(value, str):
        value = str(value)
    value = value.rstrip(' ')
    return value if value != '' else NULL


def build(sheet, table, wb, out_dir, bin_dir):
    with open(os.path.join(bin_dir, table), 'rb') as f:
        fields, types, widths, scales, _ = parse_field_defs(f.read())

    ws = wb[sheet]
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip() for h in next(rows)]

    col_of = {name: i for i, name in enumerate(header)}
    missing = [h for h in header if h not in fields]
    if missing:
        sys.exit(f'ERROR: {sheet} has column(s) not in the {table} binary schema: {missing}')

    src_idx = [col_of.get(name) for name in fields]
    extra = [name for name, i in zip(fields, src_idx) if i is None]

    out_path = os.path.join(out_dir, table + '.csv')
    os.makedirs(out_dir, exist_ok=True)
    n = 0
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f, delimiter='|')
        w.writerow(fields)
        for row in rows:
            if all(v is None for v in row):
                continue
            out = []
            for j, i in enumerate(src_idx):
                if i is None:
                    out.append(NULL)
                else:
                    out.append(norm_cell(row[i], types[j], widths[j], scales[j]))
            w.writerow(out)
            n += 1

    print(f'  {table}: {n} row(s), {len(fields)} field(s) '
          f'({len(header)} from sheet, {len(extra)} M3-only left unset: {", ".join(extra)})')
    print(f'    -> {out_path}')
    return n


def main():
    xlsx, out_dir, bin_dir = sys.argv[1], sys.argv[2], sys.argv[3]
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    print(f'Reading {xlsx}')
    for sheet, table in TABS.items():
        build(sheet, table, wb, out_dir, bin_dir)


if __name__ == '__main__':
    main()
