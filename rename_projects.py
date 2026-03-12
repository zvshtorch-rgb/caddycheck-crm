"""Standardize project names in Invoice details sheet."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import openpyxl

DATA_FILE = Path(__file__).parent / "data" / "CaddyCheckProjectsInfo.xlsx"

# {old_name_lower: canonical_name}  — case-insensitive match, exact canonical output
RENAMES = {
    # Proxy Delhaize Denderleew
    "proxy delhaize denderleeuw":           "Proxy Delhaize Denderleew",

    # Spar Dadizele
    "spar dadizele":                        "Spar Dadizele",
    "spar dadizele (2nd)":                  "Spar Dadizele",

    # Proxy Delhaize Denderhoutem
    "proxy delhaize denderhoutem (back)":   "Proxy Delhaize Denderhoutem",
    "proxy delhaize denderhoutem (top)":    "Proxy Delhaize Denderhoutem",

    # Plus Simpelveld Schouteten
    "plus simpelveld schouteten":           "Plus Simpelveld Schouteten",

    # Plus Klazienaveen Fischer
    "plus klazienaveen":                    "Plus Klazienaveen Fischer",
    "plus klazienaveen fischer":            "Plus Klazienaveen Fischer",

    # AH Merksem MAATJES (strip _x000D_ / \r)
    "ah merksem maatjes":                   "AH Merksem MAATJES",

    # AD Aartselar
    "ad aartselar (1 top)":                 "AD Aartselar",
    "ad aartselar (4 back)":                "AD Aartselar",
    "ad aartselar (4 top)":                 "AD Aartselar",
    "ad delhaize aartselaar":               "AD Aartselar",

    # Coop Bert Stuut
    "coop bert stuut (back)":              "Coop Bert Stuut",
    "coop bert stuut (top)":               "Coop Bert Stuut",

    # Jumbo Eindhoven Boschdijk
    "jumbo eindhoven bosdijk":             "Jumbo Eindhoven Boschdijk",
    "jumbo eindhoven bosdijk (from side to td)": "Jumbo Eindhoven Boschdijk",
    "jumbo eindhoven boschdijk":           "Jumbo Eindhoven Boschdijk",
}

def normalize(val):
    return str(val or "").strip().replace("\r", "").replace("_x000d_", "").lower()

wb = openpyxl.load_workbook(DATA_FILE)

total = 0
for sheet_name in ["Invoice details", "Projects overview"]:
    if sheet_name not in wb.sheetnames:
        continue
    ws = wb[sheet_name]
    rows_list = list(ws.iter_rows(values_only=False))
    header = [str(c.value).strip().lower() if c.value else "" for c in rows_list[0]]
    proj_col = next((i+1 for i, h in enumerate(header) if "project" in h and "name" in h), None)
    if proj_col is None:
        proj_col = next((i+1 for i, h in enumerate(header) if "project" in h), None)
    if proj_col is None:
        print(f"  Skipping {sheet_name}: no project column found")
        continue
    updated = 0
    for row in rows_list[1:]:
        cell = row[proj_col - 1]
        if not cell.value:
            continue
        raw = str(cell.value).strip().replace("\r", "").replace("_x000d_", "")
        key = raw.lower()
        if key in RENAMES:
            cell.value = RENAMES[key]
            updated += 1
        elif raw.lower().startswith("ad delhaize "):
            # Strip "Delhaize " from "AD Delhaize X" → "AD X"
            new_name = "AD " + raw[len("AD Delhaize "):]
            cell.value = new_name
            updated += 1
        elif raw.lower().startswith("edeka ") and " - " in raw:
            # "Edeka Vogel - Dueren" → "Edeka Vogel Dueren"
            new_name = raw.replace(" - ", " ")
            cell.value = new_name
            updated += 1
    print(f"{sheet_name}: {updated} cells renamed.")
    total += updated

wb.save(DATA_FILE)
print(f"\nTotal: {total} cells renamed. Saved.")
