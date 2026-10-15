# -*- coding: utf-8 -*-
"""
fix_cell24.py
Fix Cell 24 (weegtabel export):
  - 'gemeente' does not exist; use 'regioprovincie' (from rename cell)
  - 'opleiding 3 niveaus' does not exist; use 'Opleiding (3 niveaus)'
"""
import json

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


# Find the weegtabel export cell
for idx, cell in enumerate(cells):
    if cell["cell_type"] == "code":
        src = "".join(cell["source"])
        if "weegtabel_export" in src and "gemeente" in src:
            print(f"Found Cell {idx} (weegtabel export)")
            new_src = """\
import pandas as pd
weight_col = "weegfactor"

# Gebruik regioprovincie (= hernoemde Provincie kolom) i.p.v. gemeente
# en Opleiding (3 niveaus) (hernoemde opleiding3 kolom)
export_cols = ['regioprovincie', 'leeftijd categorie', 'Opleiding (3 niveaus)', weight_col]
avail = [c for c in export_cols if c in df.columns]
missing = [c for c in export_cols if c not in df.columns]
if missing:
    print(f"Let op: kolommen niet gevonden in df: {missing}")

export_df = (
    df[avail]
    .dropna(subset=avail[:-1])
    .drop_duplicates(subset=avail[:-1])
    .sort_values(avail[:-1])
    .rename(columns={
        'regioprovincie':       'Provincie',
        'leeftijd categorie':   'Leeftijd',
        'Opleiding (3 niveaus)':'Opleiding',
        weight_col:             'Weegfactor',
    })
    .reset_index(drop=True)
)
export_df.to_excel("weegtabel_export.xlsx", index=False)
print(f"Weegtabel geëxporteerd: {len(export_df)} rijen naar weegtabel_export.xlsx")
"""
            update(idx, new_src)
            print(f"  Fixed Cell {idx}")
            break

# Verify
remaining = [i for i, c in enumerate(cells)
             if c["cell_type"] == "code" and "gemeente" in "".join(c["source"])
             and "weegtabel" in "".join(c["source"])]
if remaining:
    print(f"WARNING: Cell(s) still have 'gemeente' in weegtabel context: {remaining}")
else:
    print("Verification: no remaining gemeente references in weegtabel cell")

with open(NB, "w", encoding="utf-8") as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print(f"Done. Total cells: {len(cells)}")
