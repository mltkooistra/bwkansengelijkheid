# -*- coding: utf-8 -*-
"""
test_new_cells.py
Runs ALL notebook cells and reports which ones succeed/fail.
"""
import json, ast, traceback
import matplotlib
matplotlib.use("Agg")

NB = "analyses.ipynb"
with open(NB, "r", encoding="utf-8") as f:
    nb = json.load(f)

cells = nb["cells"]
print(f"Total cells: {len(cells)}")

# Skip cells that only display (head(), etc.) or are markdown
SKIP_CELLS = {3, 5}  # .head() calls

ns = {}
results = {}

for idx, cell in enumerate(cells):
    if cell["cell_type"] != "code":
        results[idx] = "MARKDOWN"
        continue
    src = "".join(cell["source"])
    if not src.strip():
        results[idx] = "EMPTY"
        continue
    if idx in SKIP_CELLS:
        results[idx] = "SKIPPED"
        continue
    try:
        exec(compile(src, f"<cell {idx}>", "exec"), ns)
        results[idx] = "OK"
    except Exception as e:
        results[idx] = f"ERROR: {type(e).__name__}: {e}"

# Summary
print("\n=== RESULTS ===")
for idx, res in results.items():
    ct = cells[idx]["cell_type"][:4]
    src_preview = "".join(cells[idx]["source"])[:50].replace("\n", " ")
    status = res if res in ("OK", "MARKDOWN", "EMPTY", "SKIPPED") else res
    marker = "OK" if res == "OK" else ("--" if res in ("MARKDOWN", "EMPTY", "SKIPPED") else "!!")
    print(f"  [{marker}] Cell {idx:2d} ({ct}): {src_preview[:45]}")
    if res.startswith("ERROR"):
        print(f"          {res}")

# Count
ok   = sum(1 for r in results.values() if r == "OK")
err  = sum(1 for r in results.values() if r.startswith("ERROR"))
skip = sum(1 for r in results.values() if r in ("MARKDOWN", "EMPTY", "SKIPPED"))
print(f"\nOK: {ok}  |  ERRORS: {err}  |  SKIPPED/MARKDOWN: {skip}")
