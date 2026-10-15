# -*- coding: utf-8 -*-
"""Fix the garbled column name in Cell 14 of analyses.ipynb."""
import json, re

NB = "analyses.ipynb"

with open(NB, "r", encoding="utf-8") as f:
    nb = json.load(f)

cells = nb["cells"]
src14 = "".join(cells[14]["source"])

# The column name uses ë (U+00EB). Find any form near 'categorie' + non-word + 'n'
# and replace with the correct UTF-8 string.
correct = "Inkomen in 3 categorieën"  # ë = U+00EB

# Try to find and replace the broken form
# From hex check: bytes c3ab = ë in UTF-8, so after JSON decode it IS ë
# But the check via pipe failed because of stdin codec. Let's verify directly:
if correct in src14:
    print("Cell 14 column name already correct – no fix needed.")
else:
    # Replace any variant like categorie<any non-ascii>n or categorie?n
    fixed = re.sub(r"Inkomen in 3 categorie[^\w\s\"']{0,4}n", correct, src14)
    if fixed == src14:
        print("WARNING: could not find pattern to fix in Cell 14. Printing context:")
        for line in src14.splitlines():
            if "categorie" in line.lower():
                print(repr(line))
    else:
        cells[14]["source"] = fixed.splitlines(True)
        if cells[14]["source"] and cells[14]["source"][-1].endswith("\n"):
            cells[14]["source"][-1] = cells[14]["source"][-1].rstrip("\n")
        print(f"Fixed: {correct!r} inserted in Cell 14.")
        with open(NB, "w", encoding="utf-8") as f:
            json.dump(nb, f, ensure_ascii=False, indent=1)
        print("Notebook saved.")

# Final verification
with open(NB, "r", encoding="utf-8") as f:
    nb2 = json.load(f)
src14v = "".join(nb2["cells"][14]["source"])
print("Verification - contains correct column name:", correct in src14v)
print("Verification - contains buurt save filename:",
      "veiligheid_buurt_gewogen.png" in src14v)
