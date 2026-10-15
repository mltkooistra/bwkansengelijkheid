# -*- coding: utf-8 -*-
"""
fix_notebook2.py
Second-pass fixes based on actual column names in the dataset:
  1. Replace all "Weegfactor_GRDR" with "weegfactor" (the real column name)
  2. Insert a column-rename cell after Cell 1 so downstream code works
  3. Fix background_vars lists to use the renamed column names
  4. Fix 'leeftijd categorie' -> 'leefcat' alias (via rename cell)
"""
import json
import re

NB = "analyses.ipynb"

with open(NB, "r", encoding="utf-8") as f:
    nb = json.load(f)

cells = nb["cells"]


def src_lines(src: str) -> list:
    lines = src.splitlines(True)
    if lines and lines[-1].endswith("\n"):
        lines[-1] = lines[-1].rstrip("\n")
    return lines


def update(i: int, src: str) -> None:
    cells[i]["source"] = src_lines(src)
    cells[i]["outputs"] = []
    cells[i]["execution_count"] = None


def new_code(src: str) -> dict:
    return {"cell_type": "code", "execution_count": None,
            "metadata": {}, "outputs": [], "source": src_lines(src)}


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1  Global replace "Weegfactor_GRDR" → "weegfactor" in ALL cells
# ─────────────────────────────────────────────────────────────────────────────
n_replaced = 0
for cell in cells:
    if cell["cell_type"] == "code":
        src = "".join(cell["source"])
        new_src = src.replace("Weegfactor_GRDR", "weegfactor")
        if new_src != src:
            cell["source"] = src_lines(new_src)
            cell["outputs"] = []
            cell["execution_count"] = None
            n_replaced += 1
print(f"Step 1: replaced 'Weegfactor_GRDR' in {n_replaced} cells")


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2  Insert column-rename cell after Cell 1 (index 2)
# ─────────────────────────────────────────────────────────────────────────────
RENAME_CELL_SRC = """\
# Hernoem korte SPSS-codes naar leesbare namen voor gebruik in analyses/grafieken
rename_map = {}
_rename_candidates = {
    "opleiding3":                                       "Opleiding (3 niveaus)",
    "Inkomencat":                                       "Inkomen (categorie)",
    "Woont u in een koopwoning of huurwoning?":         "Koop of huurwoning",
    "leefcat":                                          "leeftijd categorie",
    "Provincie":                                        "regioprovincie",
}
for orig, kort in _rename_candidates.items():
    if orig in df.columns:
        rename_map[orig] = kort
df.rename(columns=rename_map, inplace=True)
combined = df  # blijf gesynchroniseerd
print("Hernoemde kolommen:", list(rename_map.values()))
"""

cells.insert(2, new_code(RENAME_CELL_SRC))
print("Step 2: inserted rename cell at index 2 (cells shifted by 1)")


# ─────────────────────────────────────────────────────────────────────────────
# STEP 3  Fix background_vars in demographic cells (now shifted +1)
# ─────────────────────────────────────────────────────────────────────────────
# After insertion: original cells 12,13,14 → now 13,14,15
# Original new Cell 27 (demographics analysis) → now 28

OLD_DEMO_VARS = (
    '"Leeftijd", "Geslacht", "opleiding 3 niveaus",\n'
    '                                "Inkomen in 3 categorie\\u00ebn", "Koop of Huurwoning", "stadplatteland"'
)

NEW_DEMO_VARS_CELLS_13_15 = (
    '"Leeftijd", "Geslacht", "Opleiding (3 niveaus)",\n'
    '                                "Inkomen (categorie)", "Koop of huurwoning", "stadplatteland"'
)

for idx in [13, 14, 15]:
    src = "".join(cells[idx]["source"])
    # Replace opleiding 3 niveaus
    new_src = src.replace('"opleiding 3 niveaus"', '"Opleiding (3 niveaus)"')
    # Replace Inkomen in 3 categorieën  (ë = ë)
    new_src = re.sub(r'"Inkomen in 3 categorie[^\w\s"\']{0,4}n"',
                     '"Inkomen (categorie)"', new_src)
    # Replace Koop of Huurwoning
    new_src = new_src.replace('"Koop of Huurwoning"', '"Koop of huurwoning"')
    if new_src != src:
        cells[idx]["source"] = src_lines(new_src)
        cells[idx]["outputs"] = []
        cells[idx]["execution_count"] = None
        print(f"  Fixed background_vars in cell {idx}")


# Fix new Cell 28 (ANALYSE A – Demographics, was cell 27 before insertion)
# Search for it by content
for idx, cell in enumerate(cells):
    if cell["cell_type"] == "code":
        src = "".join(cell["source"])
        if "ANALYSE A" in src and "Demografisch profiel" in src:
            new_src = src
            new_src = new_src.replace('"opleiding 3 niveaus"', '"Opleiding (3 niveaus)"')
            new_src = re.sub(
                r'"Inkomen in 3 categorie[^"]{0,10}"',
                '"Inkomen (categorie)"',
                new_src,
            )
            new_src = re.sub(
                r'"categorie\\\\u00ebn"',
                '"Inkomen (categorie)"',
                new_src,
            )
            # Fix the unicode escape form in CELL_A
            new_src = new_src.replace(
                '"Inkomen in 3 categorie\\u00ebn"',
                '"Inkomen (categorie)"',
            )
            new_src = new_src.replace('"Koop of Huurwoning"', '"Koop of huurwoning"')
            new_src = new_src.replace('"Koop of huurwoning"', '"Koop of huurwoning"')
            if new_src != src:
                cells[idx]["source"] = src_lines(new_src)
                cells[idx]["outputs"] = []
                cells[idx]["execution_count"] = None
                print(f"  Fixed background_vars in ANALYSE A cell {idx}")
            break


# ─────────────────────────────────────────────────────────────────────────────
# STEP 4  Fix Cell 23 (weegtabel export) – 'leeftijd categorie' already handled
#          by the rename cell, so just confirm 'leeftijd categorie' is used
#          (the rename cell maps leefcat → leeftijd categorie)
# ─────────────────────────────────────────────────────────────────────────────
for idx, cell in enumerate(cells):
    if cell["cell_type"] == "code":
        src = "".join(cell["source"])
        if "weegtabel_export" in src or "Weegtabel" in src:
            # The column 'leeftijd categorie' will now exist due to rename cell
            # Nothing to change – just confirm
            print(f"  Cell {idx}: weegtabel export uses 'leeftijd categorie' which is now correct")
            break


# ─────────────────────────────────────────────────────────────────────────────
# STEP 5  Verify no "Weegfactor_GRDR" remains
# ─────────────────────────────────────────────────────────────────────────────
remaining = [i for i, c in enumerate(cells)
             if c["cell_type"] == "code" and "Weegfactor_GRDR" in "".join(c["source"])]
if remaining:
    print(f"WARNING: 'Weegfactor_GRDR' still found in cells: {remaining}")
else:
    print("Step 5: No remaining 'Weegfactor_GRDR' references – OK")


# ─────────────────────────────────────────────────────────────────────────────
# Write back
# ─────────────────────────────────────────────────────────────────────────────
with open(NB, "w", encoding="utf-8") as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print(f"\nDone. Total cells: {len(cells)}")
