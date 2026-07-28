import os
import re
import pyodbc
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

# Configs
FOLDER_PATH = r"C:\\ASRaymond"
DSN = "Raymond DB"
USER = "ericp"
PASSWORD = "EPM3con531"
MAX_THREADS = 5

QUERY_NAME_PATTERN = re.compile(r"Query Name:\s*(\w+)")
SCHEMA_PATTERN = re.compile(r'Schema = .*?\[Name=\"(.*?)\"')
SOURCE_PATTERN = re.compile(r'(\w+)_View = .*?\[Name=\"(.*?)\",Kind=\"View\"\]')
FINAL_VAR_PATTERN = re.compile(r'in\s+(#?.+)$')
COLUMN_SELECT_PATTERN = re.compile(r'SelectColumns\(.*?\{(.*?)\}\)')


def extract_sql_parts(content):
    schema_match = SCHEMA_PATTERN.search(content)
    source_match = SOURCE_PATTERN.search(content)
    final_var_match = FINAL_VAR_PATTERN.search(content)

    if not schema_match or not source_match or not final_var_match:
        raise ValueError("Could not extract schema/source/final step")

    schema = schema_match.group(1)
    table = source_match.group(2)
    final_var = final_var_match.group(1).strip().replace('"', '')

    # extract full let block
    let_block = content.split('in')[0]
    let_lines = [line.strip().rstrip(',') for line in let_block.splitlines() if '=' in line]

    filters = []
    select_cols = []
    for line in let_lines:
        if final_var in line:
            if 'SelectRows' in line:
                expr_match = re.search(r'each\s*(.*)\)', line)
                if expr_match:
                    filters.append(expr_match.group(1))
            if 'SelectColumns' in line:
                col_match = COLUMN_SELECT_PATTERN.search(line)
                if col_match:
                    cols = [c.strip().strip('"') for c in col_match.group(1).split(',')]
                    select_cols = cols

    return schema, table, filters, select_cols


def transform_filters(filters):
    sql_conditions = []
    for f in filters:
        f = f.replace("and", "AND").replace("or", "OR")
        f = re.sub(r'\[([\w\d_]+)\]', r'\1', f)
        f = re.sub(r'"([^"]+)"', r"'\1'", f)
        f = f.replace('Text.Contains(', '').replace(')', '')
        if '@' in f:
            col, val = f.split(',')
            sql_conditions.append(f"{col.strip()} LIKE '%@%'")
        elif '<>' in f or '=' in f:
            sql_conditions.append(f)
    return ' AND '.join(sql_conditions)


def build_sql_query(schema, table, filters, columns):
    qualified_table = f"{schema}.{table}"
    cols = ', '.join(columns) if columns else '*'
    where_clause = f" WHERE {transform_filters(filters)}" if filters else ''
    return f"SELECT {cols} FROM {qualified_table}{where_clause}"


def parse_query_file(filepath):
    with open(filepath, 'r') as f:
        content = f.read()

    query_name_match = QUERY_NAME_PATTERN.search(content)
    if not query_name_match:
        raise ValueError(f"Missing Query Name in {filepath}")

    schema, table, filters, columns = extract_sql_parts(content)
    return {
        'query_name': query_name_match.group(1),
        'schema': schema,
        'table_name': table,
        'filters': filters,
        'columns': columns
    }


def process_query_file(filepath):
    try:
        parsed = parse_query_file(filepath)
        query_name = parsed['query_name']
        schema = parsed['schema']
        table = parsed['table_name']
        filters = parsed['filters']
        columns = parsed['columns']

        sql = build_sql_query(schema, table, filters, columns)

        conn_str = f"DSN={DSN};UID={USER};PWD={PASSWORD}"
        with pyodbc.connect(conn_str) as conn:
            df = pd.read_sql(sql, conn)

        output_path = os.path.join(FOLDER_PATH, f"{query_name}.xlsx")
        df.to_excel(output_path, index=False)

        print(f"✅ {query_name} exported to {output_path}")
    except Exception as e:
        print(f"❌ Error processing {filepath}: {e}")


def main():
    txt_files = [os.path.join(FOLDER_PATH, f) for f in os.listdir(FOLDER_PATH) if f.endswith('_powerqueries.txt')]
    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        executor.map(process_query_file, txt_files)


if __name__ == '__main__':
    main()