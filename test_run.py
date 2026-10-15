# -*- coding: utf-8 -*-
"""
test_run.py
Simulates running the key setup cells of the notebook to verify correctness.
Runs cells 0-17 (data load through situatie creation and demographics).
"""
import json, ast, sys, traceback
import matplotlib
matplotlib.use("Agg")  # non-interactive backend

NB = "analyses.ipynb"
with open(NB, "r", encoding="utf-8") as f:
    nb = json.load(f)

cells = nb["cells"]

# Run cells 0-17 (excluding display-only cells 3, 5)
skip_cells = {3, 5}  # head() cells – skip for non-interactive run
run_cells  = list(range(18))

ns = {}  # shared namespace

for idx in run_cells:
    cell = cells[idx]
    if cell["cell_type"] != "code":
        continue
    src = "".join(cell["source"])
    if not src.strip():
        continue
    if idx in skip_cells:
        print(f"Cell {idx:2d}: SKIPPED (display only)")
        continue
    try:
        exec(compile(src, f"<cell {idx}>", "exec"), ns)
        # Print key results
        if idx == 1:
            print(f"Cell {idx:2d}: df.shape = {ns.get('df', None) and ns['df'].shape}")
        elif idx == 2:
            print(f"Cell {idx:2d}: rename cell OK – new cols present: "
                  + str(['leeftijd categorie' in ns['df'].columns,
                          'Koop of huurwoning' in ns['df'].columns,
                          'Opleiding (3 niveaus)' in ns['df'].columns]))
        elif idx == 4:
            n = len(ns.get('tevreden_cols', []))
            print(f"Cell {idx:2d}: tevreden_cols = {n}")
        elif idx == 6:
            n = len(ns.get('toekomst_cols', []))
            print(f"Cell {idx:2d}: toekomst_cols = {n}")
        elif idx == 9:
            subj = ns.get('subjects', [])
            print(f"Cell {idx:2d}: subjects = {len(subj)}: {subj[:3]}...")
        elif idx == 10:
            w = ns.get('weight_col', '?')
            print(f"Cell {idx:2d}: weight_col = {w!r}, exists={w in ns['df'].columns}")
        elif idx == 12:
            s = ns.get('df', None)
            scols = [c for c in (s.columns if s is not None else []) if c.startswith('situatie_')]
            print(f"Cell {idx:2d}: situatie cols = {len(scols)}: {scols[:2]}...")
        elif idx == 13:
            print(f"Cell {idx:2d}: stacked bar OK")
        elif idx == 17:
            inv = ns.get('invloed_dummy_cols', [])
            print(f"Cell {idx:2d}: invloed_dummy_cols = {len(inv)}")
        else:
            print(f"Cell {idx:2d}: OK")
    except Exception as e:
        print(f"Cell {idx:2d}: ERROR – {type(e).__name__}: {e}")
        traceback.print_exc()

print("\nTest run complete.")
