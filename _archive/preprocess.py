import re

def split_queries(text):
    queries = []
    parts = re.split(r'(?=Query Name:)', text)
    for part in parts:
        if part.strip():
            name_match = re.search(r'Query Name:\s*(\w+)', part)
            name = name_match.group(1) if name_match else "Unknown"
            queries.append((name, part))
    return queries

def extract_steps(query_text):
    # Extract the let/in block
    let_in_match = re.search(r'let(.*?)in\s+(.+)', query_text, re.DOTALL)
    if not let_in_match:
        return []
    let_block = let_in_match.group(1)
    steps = []
    for line in let_block.splitlines():
        line = line.strip().rstrip(',')
        if '=' in line:
            step_name, expr = line.split('=', 1)
            steps.append((step_name.strip(), expr.strip()))
    return steps

if __name__ == "__main__":
    with open(r"C:\ASRaymond\INP35_powerqueries.txt", encoding="utf-8", errors="replace") as f:
        text = f.read()
    queries = split_queries(text)
    for name, qtext in queries:
        print(f"--- {name} ---")
        steps = extract_steps(qtext)
        for step in steps:
            print(step)
        print()