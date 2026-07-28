from pathlib import Path

# List of facilities
facilities = [
    "AEX", "ALV", "ATL", "BAK", "BCA", "BLU", "CIN", "CO", "CTO", "FON", "FW", "HT",
    "JAC", "KC", "KP", "LKC", "LKP", "LKS", "LYN", "MAN", "MD", "OH", "OK", "PHX",
    "POU", "SAC", "SED", "SLT", "SPA", "SYR", "TAL", "TAM", "UCO", "WI", "WIB",
    "WND", "WYA", "YUM", "POR", "FWP"
]

# Output folder
output_dir = Path("queries/evs")
output_dir.mkdir(parents=True, exist_ok=True)

# SQL template
sql_template = """SELECT 
    ''[MESSAGE],
    Facility[FACI],
    Item[ITNO],
    ''[STRT],
    '3'[PCTP],
    '081825'[PCDT],
    Cost[CSU1] 
FROM CTOS_TerexItems, CTOS_Facilities 
WHERE Facility = '{facility}' 
ORDER BY 1,2
"""

# Generate files
for fac in facilities:
    sql_content = sql_template.format(facility=fac)
    file_path = output_dir / f"API_EXT003MI_AddStandardCost_[{fac}].sql"
    file_path.write_text(sql_content, encoding="utf-8")

print(f"✅ Generated {len(facilities)} SQL files in {output_dir.resolve()}")