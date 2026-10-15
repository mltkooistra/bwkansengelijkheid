# -*- coding: utf-8 -*-
"""
fix_notebook3.py
Final fixes based on actual data values:
  1. Cell 4: tevredenheid values are already float64 – drop reverse_map,
             just use pd.to_numeric() directly.
  2. Cell 14: situatie_veiligheid doesn't exist – change to
             situatie_uw_lichamelijke_gezondheid.
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


# ─────────────────────────────────────────────────────────────────────────────
# FIX 1  Cell 4 – tevredenheid scores zijn al float64 (1.0-10.0)
#         De reverse_map aanpak converteert ze onterecht naar NaN.
#         Oplossing: gebruik pd.to_numeric() direct.
# ─────────────────────────────────────────────────────────────────────────────
for idx in range(len(cells)):
    src = "".join(cells[idx]["source"])
    if ("tevreden_cols" in src and "reverse_map" in src and
            "tevreden_df" in src and "meta_labels" in src):
        print(f"Found Cell {idx} (tevreden reverse_map issue)")
        new_src = """\
import pandas as pd

# Filter tevredenheid columns
tevreden_cols = [col for col in combined.columns if "hoe tevreden u bent met" in col]

# Clean up column names for the matrix (extract the subject after the last " - ")
short_names = {col: col.split(" - ")[-1] for col in tevreden_cols}

tevreden_df = combined[tevreden_cols].rename(columns=short_names)

# Tevredenheidsscores zijn al numeriek (1-10 float64) – direct omzetten
for col in tevreden_df.columns:
    tevreden_df[col] = pd.to_numeric(tevreden_df[col], errors="coerce")

# Correlation matrix
corr_matrix_tevredenheid = tevreden_df.corr()

print(f"Tevredenheid: {len(tevreden_cols)} kolommen, {tevreden_df.notna().mean().mean()*100:.1f}% niet-missing")
"""
        update(idx, new_src)
        break

# ─────────────────────────────────────────────────────────────────────────────
# FIX 2  Cell with situatie_veiligheid (doesn't exist) → gebruik gezondheid
# ─────────────────────────────────────────────────────────────────────────────
for idx in range(len(cells)):
    src = "".join(cells[idx]["source"])
    if 'situatie_leven = "situatie_veiligheid"' in src:
        print(f"Found Cell {idx} (situatie_veiligheid fix)")
        new_src = src.replace(
            'situatie_leven = "situatie_veiligheid"',
            'situatie_leven = "situatie_uw_lichamelijke_gezondheid"',
        ).replace(
            'fig.suptitle("Achtergrondkenmerken (gewogen)"',
            'fig.suptitle("Achtergrondkenmerken: lichamelijke gezondheid (gewogen)"',
        ).replace(
            'plt.savefig("descriptives_situatie_veiligheid_gewogen.png"',
            'plt.savefig("descriptives_situatie_gezondheid_gewogen.png"',
        )
        update(idx, new_src)
        break

# ─────────────────────────────────────────────────────────────────────────────
# Verify
# ─────────────────────────────────────────────────────────────────────────────
for idx in range(len(cells)):
    src = "".join(cells[idx]["source"])
    if "reverse_map" in src and "tevreden_df" in src:
        print(f"WARNING: Cell {idx} still has reverse_map for tevreden_df")
    if "situatie_veiligheid\"" in src:
        print(f"WARNING: Cell {idx} still has situatie_veiligheid")

print("All verifications passed.")

# ─────────────────────────────────────────────────────────────────────────────
# Write back
# ─────────────────────────────────────────────────────────────────────────────
with open(NB, "w", encoding="utf-8") as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print(f"Done. Total cells: {len(cells)}")
