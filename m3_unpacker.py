#!/usr/bin/env python3
"""
m3_unpacker.py — Parse M3 Grid Access binary table export files to CSV (Pipe Delimited).

Companion file TABLE_INFO (Java-serialized ArrayList) optionally provides
expected record counts per table, used to validate output.
"""

import struct, csv, sys, os, argparse

# CSV token for a field that is NOT stored on a record (bitmap bit clear).
# A plain empty cell means the field IS stored but empty (M3's LPS length-0
# "inherit" marker) — a distinct wire state. Only raw/round-trip mode (fill=
# False) emits these; m3_repacker.py reads the sentinel back to None so the
# original sparse bitmap is reproduced byte-for-byte. Keep this in sync with
# m3_repacker.NULL_SENTINEL.
NULL_SENTINEL = r'\N'

# ─────────────────────────────────────────────
# TABLE_INFO parser
# ─────────────────────────────────────────────

def parse_table_info(path):
    """Scan a Java-serialized ArrayList (TABLE_INFO) for table names and their record counts.

    The file is produced by M3 Grid Access and uses Java's ObjectOutputStream format.
    0x74 is the TC_STRING opcode; the following 2 bytes are a big-endian u16 string
    length. The 8 bytes immediately preceding the opcode encode the record count as a
    big-endian signed 64-bit integer (Java long), which is how M3 lays out the
    ArrayList entries.

    Returns a dict of {table_name: record_count}, or {} if the file is absent.
    """
    try:
        with open(path, 'rb') as f:
            data = f.read()
    except FileNotFoundError:
        return {}

    tables = {}
    i = 0
    while i < len(data) - 2:
        if data[i] == 0x74:  # Java TC_STRING opcode
            name_len = struct.unpack('>H', data[i+1:i+3])[0]
            if i + 3 + name_len <= len(data):
                try:
                    name = data[i+3:i+3+name_len].decode('ascii')
                    # M3 table names are ALL_CAPS, 4–12 chars; filter out other strings
                    # that also happen to follow a TC_STRING opcode.
                    if name.isupper() and 4 <= len(name) <= 12 and i >= 8:
                        no_recs = struct.unpack('>q', data[i-8:i])[0]
                        if 0 <= no_recs < 10_000_000:
                            tables[name] = no_recs
                except (UnicodeDecodeError, struct.error):
                    pass
        i += 1
    return tables


# Fixed byte regions of the Java-serialized ArrayList<ToolProxy$TableInfo>
# stream, captured from a real M3-exported TABLE_INFO file. HEAD covers
# STREAM_MAGIC through the ArrayList classDesc (ending at its TC_NULL
# superclass); _MID covers the first element's TC_OBJECT + TC_CLASSDESC for
# ToolProxy$TableInfo (ending at its TC_NULL superclass). Neither depends on
# entry count or table name, so later entries reuse _MID's classDesc via a
# TC_REFERENCE to _TABLEINFO_CLASSDESC_HANDLE instead of re-serializing it —
# mirroring how a real JVM ObjectOutputStream back-references a repeated
# class descriptor.
_TABLE_INFO_HEAD = (
    b'\xac\xed\x00\x05sr\x00\x13java.util.ArrayListx'
    b'\x81\xd2\x1d\x99\xc7a\x9d\x03\x00\x01I\x00\x04sizexp'
)
_TABLE_INFO_MID = (
    b'sr\x001gridaccess.client.tools.proxy.ToolProxy$TableInfo'
    b'\x00\x00\x00\x00\x00\x00\x00\x01\x02\x00\x02J\x00\tnoRecordsL\x00\ttableName'
    b't\x00\x12Ljava/lang/String;xp'
)
_TABLEINFO_CLASSDESC_HANDLE = 0x7E0002


def write_table_info(entries):
    """Build a TABLE_INFO file (Java-serialized ArrayList<TableInfo>) from
    an iterable of (table_name, record_count) pairs. Inverse of
    parse_table_info(): write_table_info(parse_table_info(p).items()) round
    trips to a stream a real M3 client can read, though not necessarily the
    exact same bytes as the original (e.g. re-ordering, or JVM-side capacity
    choices) — verified byte-identical for the single-entry case, which
    covers the normal single-table repack workflow.
    """
    entries = list(entries)
    n = len(entries)
    out = bytearray(_TABLE_INFO_HEAD)
    out += struct.pack('>i', n)                            # ArrayList.size
    out += b'\x77\x04' + struct.pack('>i', n)               # writeObject() capacity
    for idx, (name, count) in enumerate(entries):
        if idx == 0:
            out += _TABLE_INFO_MID
        else:
            out += b'\x73\x71' + struct.pack('>i', _TABLEINFO_CLASSDESC_HANDLE)
        out += struct.pack('>q', count)
        name_bytes = name.encode('ascii')
        out += b'\x74' + struct.pack('>H', len(name_bytes)) + name_bytes
    out += b'\x78'                                          # TC_ENDBLOCKDATA
    return bytes(out)


# ─────────────────────────────────────────────
# Field definition parser
# ─────────────────────────────────────────────

def parse_field_defs(data):
    """Returns (field_names, field_types, field_widths, field_scales, header_length).
    field_types are JDBC type codes as strings, e.g. '4'=INTEGER, '12'=VARCHAR,
    '2'=NUMERIC (decimal), '-5'=BIGINT.
    field_widths are the display widths from the field definition (e.g. 8 for a date).
    field_scales are the number of decimal places (the 4th component of each
    field def, e.g. 2 in "2;MMXIMP;5;2"). This is what tells a decimal field
    (scale > 0, LPS-encoded) apart from an integer/date field (scale 0, which,
    when narrow enough, is stored as a raw 4-byte int). See _is_short_decimal.
    """
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
                # 4th component is the decimal scale; older/edge defs may omit it.
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


# JDBC type → display default for a field that's bit-clear (not stored on the record)
_NUMERIC_TYPES = {'4', '-5', '5', '-6', '7', '6', '8', '2', '3'}

def default_for_type(type_code):
    """Default value M3 displays for a field that's not stored on a record."""
    if type_code in _NUMERIC_TYPES:
        return '0'
    return ''


# Whole-number fields (scale 0) whose display width fits in a signed 32-bit int
# (max 999,999,999) are stored in the binary as a raw 4-byte big-endian int — no
# LPS length prefix. Any field with decimal places (scale > 0) is stored as an
# LPS string (e.g. "1.44"), regardless of width, and so are wider whole numbers
# (width >= 10). Ignoring the scale — treating every narrow numeric as a 4-byte
# int — misreads decimal fields and drifts the whole record; that was the bug.
_SHORT_DECIMAL_MAX_WIDTH = 9


def _is_short_decimal(type_code, width, scale=0):
    return (
        type_code in _NUMERIC_TYPES
        and (scale or 0) == 0
        and width <= _SHORT_DECIMAL_MAX_WIDTH
    )


# ─────────────────────────────────────────────
# LPS helpers
# ─────────────────────────────────────────────

MAX_FIELD_LEN = 1048576   # Increased to 1MB to safely handle large memo fields

def read_lps(data, pos):
    """Try to read a length-prefixed string at pos. Returns (value, new_pos) or None."""
    if pos + 4 > len(data):
        return None
        
    length = struct.unpack('>I', data[pos:pos+4])[0]
    
    if length > MAX_FIELD_LEN or pos + 4 + length > len(data):
        return None
        
    # M3 uses a length of 0 for empty strings.
    if length == 0:
        return '', pos + 4
        
    raw = data[pos+4:pos+4+length]
    
    # Inter-record gaps contain binary nulls. Legitimate M3 text fields do not.
    if b'\x00' in raw:
        return None
        
    try:
        # Standard UTF-8 decoding (supports all characters, newlines, and symbols)
        return raw.decode('utf-8'), pos + 4 + length
    except UnicodeDecodeError:
        try:
            # Fallback to CP1252 for older legacy European M3 environments
            return raw.decode('cp1252'), pos + 4 + length
        except UnicodeDecodeError:
            return None


# ─────────────────────────────────────────────
# Bitmap-style parser (M3 Grid Access native format)
#
# Each record is laid out as:
#   [4 bytes]  record block size (big-endian u32) — counts everything that
#              follows for this record (bitmap + all present-field LPS entries)
#   [K bytes]  nullability bitmap, K = ceil(n_fields / 8), MSB-first.
#              Bit i (counting from the MSB of byte i//8) = 1 means field i
#              is present; 0 means it's null/empty and is OMITTED from the
#              binary entirely.
#   [...]      For each present field:
#              - String / wide-decimal (width ≥ 10): 4-byte big-endian u32 length
#                prefix followed by UTF-8 bytes of the value.
#              - Short decimal (numeric type, width ≤ 9): raw 4-byte big-endian
#                signed int — no length prefix.
# ─────────────────────────────────────────────

def parse_records_bitmap_style(data_section, n_fields, field_types=None,
                               field_widths=None, field_scales=None,
                               expected_count=None, fill=True):
    """
    Parse records using M3's bitmap layout.

    String fields and wide decimal fields (width ≥ 10) use LPS encoding:
      - Bit set + length > 0 : explicit value
      - Bit set + length = 0 : "inherit" — when fill=True, carry forward the
                               last value seen for this column across records.
                               When fill=False, emit '' (distinct from None —
                               the field IS stored on the wire, just empty).
      - Bit clear            : field not stored. When fill=True, emit the
                               type's default ("0" for numeric, "" for string).
                               When fill=False, emit None — this is what lets
                               m3_repacker_db.py reconstruct the original
                               bitmap byte-for-byte (see encode_bitmap_record).

    Short decimal fields (numeric type, width ≤ 9) use raw 4-byte big-endian
    int encoding — no LPS length prefix. The integer value IS the field value.
    """
    records = []
    bitmap_len = (n_fields + 7) // 8
    pos = 0

    # Per-column carry-forward state (only used when fill=True)
    last_value = [''] * n_fields
    if field_types and fill:
        for i, t in enumerate(field_types):
            last_value[i] = default_for_type(t)

    # Pre-compute per-field storage class
    short_decimal = [False] * n_fields
    if field_types and field_widths:
        for i in range(n_fields):
            t = field_types[i] if i < len(field_types) else ''
            w = field_widths[i] if i < len(field_widths) else 0
            s = field_scales[i] if field_scales and i < len(field_scales) else 0
            short_decimal[i] = _is_short_decimal(t, w, s)

    while pos + 4 <= len(data_section):
        rec_size = struct.unpack('>I', data_section[pos:pos+4])[0]
        if rec_size == 0 or pos + 4 + rec_size > len(data_section):
            break

        rec_end = pos + 4 + rec_size
        bitmap  = data_section[pos+4:pos+4+bitmap_len]
        if len(bitmap) < bitmap_len:
            break

        read_pos = pos + 4 + bitmap_len
        values = []
        ok = True

        for i in range(n_fields):
            byte_idx = i // 8
            bit_idx  = 7 - (i % 8)
            present  = (bitmap[byte_idx] >> bit_idx) & 1
            ftype    = field_types[i] if field_types and i < len(field_types) else ''

            if not present:
                # Field not stored on this record.
                values.append(last_value[i] if fill else None)
                continue

            if short_decimal[i]:
                # Raw 4-byte big-endian signed int — no LPS prefix.
                if read_pos + 4 > rec_end:
                    ok = False
                    break
                int_val = struct.unpack('>i', data_section[read_pos:read_pos+4])[0]
                val = str(int_val)
                values.append(val)
                if fill:
                    last_value[i] = val
                read_pos += 4
                continue

            if read_pos + 4 > rec_end:
                ok = False
                break
            length = struct.unpack('>I', data_section[read_pos:read_pos+4])[0]
            if length > MAX_FIELD_LEN or read_pos + 4 + length > rec_end:
                ok = False
                break

            if length == 0:
                # "Inherit" marker: bit set but no value bytes.
                # Carry forward the last value we've seen for this column.
                values.append(last_value[i] if fill else '')
                read_pos += 4
                continue

            raw = data_section[read_pos+4:read_pos+4+length]
            try:
                val = raw.decode('utf-8')
            except UnicodeDecodeError:
                try:
                    val = raw.decode('cp1252')
                except UnicodeDecodeError:
                    val = raw.decode('utf-8', errors='replace')
            values.append(val)
            if fill:
                last_value[i] = val   # explicit value updates the carry-forward state
            read_pos += 4 + length

        if ok:
            records.append(values)

        # Jump to the next record using the block size header (don't trust
        # read_pos, so we stay aligned even if a record has trailing padding).
        pos = rec_end

        if expected_count and len(records) >= expected_count:
            break

    return records


# ─────────────────────────────────────────────
# PPP363-style parser (0xFF 0xE0 separator)
# ─────────────────────────────────────────────

def parse_records_ffe0_style(data_section, n_fields):
    """Parse records using the PPP363-style 0xFF 0xE0 separator format.

    Some older M3 export variants delimit records with the two-byte marker 0xFF 0xE0
    instead of using the bitmap layout. Each chunk between separators is a single
    record: a sequence of length-prefixed strings (LPS), followed by a 4-byte
    big-endian size hint that is a property of the separator and not a field value.
    """
    RECORD_SEP = b'\xff\xe0'
    chunks = data_section.split(RECORD_SEP)
    records = []

    for i, chunk in enumerate(chunks):
        # Chunk 0 contains the preamble and the first length hint, skip it.
        if i == 0 or not chunk:
            continue

        # Each inter-record chunk carries a 4-byte size hint at the end;
        # strip it so we don't mistake it for a field value.
        record_bytes = chunk[:-4] if i < len(chunks) - 1 else chunk
        values = []
        pos = 0
        
        while pos < len(record_bytes):
            res = read_lps(record_bytes, pos)
            if res is None:
                break
            val, new_pos = res
            values.append(val)
            pos = new_pos
            
        if values:
            while len(values) < n_fields:
                values.append('')
            records.append(values[:n_fields])
            
    return records


# ─────────────────────────────────────────────
# Format detection + dispatch
# ─────────────────────────────────────────────

def detect_format(data_section):
    """Return 'ffe0' if the data starts with PPP363-style separators, else 'gap' (bitmap)."""
    if b'\xff\xe0' in data_section[:500]:
        return 'ffe0'
    return 'gap'


# ─────────────────────────────────────────────
# Top-level file parser
# ─────────────────────────────────────────────

def parse_file(input_path, output_path=None, table_info=None, verbose=True, fill=True):
    """Parse a single M3 binary export file and write a pipe-delimited CSV.

    Reads field definitions from the file header, auto-detects the record encoding
    (bitmap vs. 0xFF 0xE0), dispatches to the appropriate parser, then writes the
    result. Returns True on success, False if the file was skipped or errored.

    Args:
        input_path:  Path to the M3 binary table export file.
        output_path: Destination CSV path; defaults to input_path with .csv extension.
        table_info:  Dict of {table_name: expected_record_count} from TABLE_INFO.
        verbose:     Print per-file status messages when True.
        fill:        When True, carry forward inherited values and fill numeric defaults.
    """
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + '.csv'

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

    n = len(fields)
    data_section = data[data_section_start:]
    table_name   = os.path.basename(input_path)
    expected     = (table_info or {}).get(table_name)
    fmt          = detect_format(data_section)

    # Fast-path skip: TABLE_INFO tells us this table has 0 records.
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

    # When fill=True every numeric field gets a "0" default, so rows are never
    # all-empty by construction. When fill=False the original behaviour applies.
    if not fill:
        records = [r for r in records if any(v for v in r)]

    if fmt == "gap" and expected and len(records) > expected:
        records = records[:expected]

    # Skip empty tables — don't write a header-only CSV.
    if not records:
        if verbose: print(f"  SKIP  {table_name}  —  0 records parsed")
        return False

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter='|')
        writer.writerow(fields)
        # In raw mode (fill=False) an absent field comes back as None; write it
        # as the NULL sentinel so it stays distinct from a present-but-empty ''
        # cell. In fill mode there are no None cells, so this is a no-op.
        writer.writerows(
            [NULL_SENTINEL if v is None else v for v in row]
            for row in records
        )

    count_note = f" (expected {expected})" if expected else ""
    if verbose:
        print(f"  OK    {table_name}  →  {os.path.basename(output_path)}  ({len(records)} record(s){count_note}, {n} field(s))")
    return True


# ─────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Parse M3 Grid Access binary table export files to pipe-delimited CSV.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "python3  m3_unpacker.py -i ./cmp200_tables/ -o ./csv_output/\n"
            "python3  m3_unpacker.py --input TABLE.bin --output TABLE.csv\n"
            "python3  m3_unpacker.py -i /path/to/input_dir           (writes CSVs next to inputs)\n"
        ),
    )
    parser.add_argument(
        '-i', '--input', dest='input', required=True,
        help="Input file OR directory containing M3 binary table exports."
    )
    parser.add_argument(
        '-o', '--output', dest='output', default=None,
        help="Output CSV file (when --input is a file) or output directory "
             "(when --input is a directory). Defaults to the input location."
    )
    parser.add_argument(
        '-q', '--quiet', action='store_true',
        help="Suppress per-file progress messages."
    )
    parser.add_argument(
        '--raw', action='store_true',
        help="Emit raw binary contents only — don't fill in inherited values "
             "from previous rows or type defaults (0 for numeric). Useful for "
             "byte-perfect round-trip with m3_repacker.py."
    )

    # When run with no arguments, show the full help (with examples) instead
    # of argparse's terse "error: the following arguments are required" line.
    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(1)

    args = parser.parse_args()
    input_arg  = args.input
    output_arg = args.output
    verbose    = not args.quiet
    fill       = not args.raw

    if not os.path.exists(input_arg):
        print(f"ERROR: input path does not exist: {input_arg}", file=sys.stderr)
        sys.exit(2)

    if os.path.isdir(input_arg):
        ti_path = os.path.join(input_arg, 'TABLE_INFO')
    else:
        ti_path = os.path.join(os.path.dirname(os.path.abspath(input_arg)), 'TABLE_INFO')

    table_info = parse_table_info(ti_path)
    if table_info and verbose:
        print(f"TABLE_INFO loaded: {table_info}")

    if os.path.isdir(input_arg):
        out_dir = output_arg or input_arg
        os.makedirs(out_dir, exist_ok=True)
        files = sorted(
            f for f in os.listdir(input_arg)
            if os.path.isfile(os.path.join(input_arg, f))
            and not f.startswith('.')
            and f not in ('TABLE_INFO',)
            and not f.endswith('.csv')
        )
        if verbose:
            print(f"Batch: {len(files)} file(s) in '{input_arg}'  →  '{out_dir}'\n")
        ok = err = 0
        for fname in files:
            if parse_file(
                os.path.join(input_arg, fname),
                os.path.join(out_dir, fname + '.csv'),
                table_info=table_info,
                verbose=verbose,
                fill=fill,
            ):
                ok += 1
            else:
                err += 1
        if verbose:
            print(f"\nDone: {ok} succeeded, {err} skipped/errored.")
    else:
        parse_file(input_arg, output_arg, table_info=table_info, verbose=verbose, fill=fill)

if __name__ == '__main__':
    main()