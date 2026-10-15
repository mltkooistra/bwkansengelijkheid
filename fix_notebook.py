"""
fix_notebook.py
Fixes bugs in analyses.ipynb and adds new analyses + Excel export.
Run from the same folder as analyses.ipynb.
"""

import json
import re

NB_PATH = "analyses.ipynb"

with open(NB_PATH, "r", encoding="utf-8") as f:
    nb = json.load(f)

cells = nb["cells"]


def src_to_lines(src: str) -> list:
    """Split source string into Jupyter line list (all but last end with \\n)."""
    lines = src.splitlines(keepends=True)
    # Ensure last line has no trailing newline (Jupyter convention)
    if lines and lines[-1].endswith("\n"):
        lines[-1] = lines[-1].rstrip("\n")
    return lines


def update(i: int, src: str) -> None:
    cells[i]["source"] = src_to_lines(src)
    cells[i]["outputs"] = []
    cells[i]["execution_count"] = None


def new_cell(src: str) -> dict:
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": src_to_lines(src),
    }


def new_md(src: str) -> dict:
    return {
        "cell_type": "markdown",
        "metadata": {},
        "source": src_to_lines(src),
    }


# ─────────────────────────────────────────────────────────────────────────────
# FIX 1  Cell 1 – sync df = combined after filter
# ─────────────────────────────────────────────────────────────────────────────
update(1, """\
combined = df.copy()

combined = combined[combined['Reactiestatus'] != 'Niet geantwoord']
df = combined  # sync: zorg dat df en combined dezelfde gefilterde dataset zijn
print(f"\\nCombined shape na Reactiestatus filter: {combined.shape}")
""")

# ─────────────────────────────────────────────────────────────────────────────
# FIX 2  Cell 3 – paths[0] → path
# ─────────────────────────────────────────────────────────────────────────────
src3 = "".join(cells[3]["source"])
src3 = src3.replace(
    "pyreadstat.read_sav(paths[0], metadataonly=True)",
    "pyreadstat.read_sav(path, metadataonly=True)",
)
update(3, src3)

# ─────────────────────────────────────────────────────────────────────────────
# FIX 3  Cell 4 – add () to .head
# ─────────────────────────────────────────────────────────────────────────────
update(4, "combined[tevreden_cols].head()")

# ─────────────────────────────────────────────────────────────────────────────
# FIX 4  Cell 10 – weight_col lowercase → Weegfactor_GRDR
# ─────────────────────────────────────────────────────────────────────────────
src10 = "".join(cells[10]["source"])
src10 = src10.replace("weight_col = 'weegfactor'", "weight_col = 'Weegfactor_GRDR'")
update(10, src10)

# ─────────────────────────────────────────────────────────────────────────────
# FIX 5  Cell 13 – save filename overwriting Cell 12's output
# ─────────────────────────────────────────────────────────────────────────────
src13 = "".join(cells[13]["source"])
src13 = src13.replace(
    'plt.savefig("descriptives_situatie_leven_gewogen.png"',
    'plt.savefig("descriptives_situatie_veiligheid_gewogen.png"',
)
update(13, src13)

# ─────────────────────────────────────────────────────────────────────────────
# FIX 6  Cell 14 – garbled ë in column name + wrong save filename
# ─────────────────────────────────────────────────────────────────────────────
src14 = "".join(cells[14]["source"])
# Fix any encoding-garbled variant of 'categorieën'
src14 = re.sub(r"categorie[^\w\s]{1,4}n", "categorieën", src14)
src14 = src14.replace(
    'plt.savefig("descriptives_situatie_veiligheid_gewogen.png"',
    'plt.savefig("descriptives_situatie_veiligheid_buurt_gewogen.png"',
)
update(14, src14)

# ─────────────────────────────────────────────────────────────────────────────
# NEW CELL A – Demografisch profiel per situatiegroep (leven in het algemeen)
# ─────────────────────────────────────────────────────────────────────────────
CELL_A = """\
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec

# ════════════════════════════════════════════════════════════════════════════════
# ANALYSE A: Demografisch profiel per situatiegroep (leven in het algemeen)
# Vergelijk: wie valt in welke groep? (leeftijd, geslacht, opleiding, etc.)
# ════════════════════════════════════════════════════════════════════════════════

situatie_col = "situatie_uw_leven_in_het_algemeen"
weight_col   = "Weegfactor_GRDR"
cat_order    = ["tevreden_en_stabiel", "tevreden_en_pessimistisch",
                "ontevreden_en_optimistisch", "ontevreden_en_pessimistisch"]
cat_labels   = {
    "tevreden_en_stabiel":         "Tevreden & stabiel",
    "tevreden_en_pessimistisch":   "Tevreden & pessimistisch",
    "ontevreden_en_optimistisch":  "Ontevreden & optimistisch",
    "ontevreden_en_pessimistisch": "Ontevreden & pessimistisch",
}
colors_g = {
    "tevreden_en_stabiel":         "#2ecc71",
    "tevreden_en_pessimistisch":   "#f39c12",
    "ontevreden_en_optimistisch":  "#3498db",
    "ontevreden_en_pessimistisch": "#e74c3c",
}

demo_vars = [v for v in [
    "Leeftijd", "Geslacht", "opleiding 3 niveaus",
    "Inkomen in 3 categorie\\u00ebn", "Koop of Huurwoning", "stadplatteland",
] if v in df.columns]

sub_main = df[[situatie_col, weight_col] + demo_vars].dropna(subset=[situatie_col])

# ── Gewogen N per groep ──────────────────────────────────────────────────────
groep_n = {g: df.loc[df[situatie_col] == g, weight_col].sum() for g in cat_order}
total_n = sum(groep_n.values())
print("Gewogen N per situatiegroep (leven in het algemeen):")
for g in cat_order:
    pct = groep_n[g] / total_n * 100
    print(f"  {cat_labels[g]:<35s}: n={groep_n[g]:,.0f}  ({pct:.1f}%)")

# ── Bouw profiel-tabel ────────────────────────────────────────────────────────
def wq(v, w, q):
    v, w = np.asarray(v, float), np.asarray(w, float)
    ok = ~(np.isnan(v) | np.isnan(w))
    v, w = v[ok], w[ok]
    if len(v) == 0:
        return np.nan
    o = np.argsort(v); v, w = v[o], w[o]
    cw = (np.cumsum(w) - 0.5 * w) / w.sum()
    return float(np.interp(q, cw, v))

profile_rows = []
for var in demo_vars:
    sub = sub_main[[situatie_col, var, weight_col]].dropna(subset=[var])
    if var == "Leeftijd":
        for g in cat_order:
            grp = sub[sub[situatie_col] == g]
            if len(grp) == 0:
                continue
            wavg = np.average(grp["Leeftijd"].astype(float),
                              weights=grp[weight_col].astype(float))
            wmed = wq(grp["Leeftijd"].values, grp[weight_col].values, 0.5)
            profile_rows.append({
                "Variabele": "Leeftijd", "Categorie": "Gemiddelde",
                "Groep": g, "Waarde": round(wavg, 1),
            })
            profile_rows.append({
                "Variabele": "Leeftijd", "Categorie": "Mediaan",
                "Groep": g, "Waarde": round(wmed, 1),
            })
    else:
        ct = sub.groupby([situatie_col, var])[weight_col].sum().unstack(fill_value=0)
        ct_pct = (ct.div(ct.sum(axis=1), axis=0) * 100).round(1)
        for cat in ct_pct.columns:
            for g in cat_order:
                val = float(ct_pct.loc[g, cat]) if g in ct_pct.index else 0.0
                profile_rows.append({
                    "Variabele": var, "Categorie": str(cat),
                    "Groep": g, "Waarde": val,
                })

profile_long = pd.DataFrame(profile_rows)
profile_wide = (
    profile_long
    .pivot_table(index=["Variabele", "Categorie"], columns="Groep", values="Waarde")
    .reindex(columns=[c for c in cat_order if c in profile_long["Groep"].unique()])
)
print("\\n=== Demografisch profiel (%, gewogen) ===")
print(profile_wide.to_string())

# ── Visualisatie ──────────────────────────────────────────────────────────────
fig, axes = plt.subplots(2, 3, figsize=(20, 13))
fig.suptitle("Demografisch profiel per situatiegroep (leven in het algemeen, gewogen)",
             fontsize=15, fontweight="bold", y=1.01)
axes_flat = axes.flatten()

# Leeftijd: gewogen boxplot
ax = axes_flat[0]
bxp_stats = []
for g in cat_order:
    grp = sub_main[sub_main[situatie_col] == g][["Leeftijd", weight_col]].dropna()
    if len(grp) == 0:
        bxp_stats.append({"med": 0, "q1": 0, "q3": 0, "whislo": 0, "whishi": 0, "fliers": []})
        continue
    v = grp["Leeftijd"].astype(float).values
    w = grp[weight_col].astype(float).values
    bxp_stats.append({
        "med": wq(v, w, 0.5), "q1": wq(v, w, 0.25), "q3": wq(v, w, 0.75),
        "whislo": float(v.min()), "whishi": float(v.max()), "fliers": [],
    })
bp = ax.bxp(bxp_stats, patch_artist=True, medianprops=dict(color="white", linewidth=2))
for patch, g in zip(bp["boxes"], cat_order):
    patch.set_facecolor(colors_g[g])
ax.set_xticks(range(1, len(cat_order) + 1))
ax.set_xticklabels([cat_labels[g] for g in cat_order], fontsize=8, rotation=15, ha="right")
ax.set_title("Leeftijd", fontweight="bold")
ax.set_ylabel("Jaar")
ax.grid(axis="y", alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)

# Categorische vars
for idx, var in enumerate([v for v in demo_vars if v != "Leeftijd"][:5]):
    ax = axes_flat[idx + 1]
    sub = sub_main[[situatie_col, var, weight_col]].dropna(subset=[var])
    ct = sub.groupby([situatie_col, var])[weight_col].sum().unstack(fill_value=0)
    ct_pct = (ct.div(ct.sum(axis=1), axis=0) * 100).reindex(cat_order).dropna(how="all")
    ct_pct.T.plot(kind="bar", ax=ax,
                  color=[colors_g[g] for g in ct_pct.index],
                  edgecolor="white", linewidth=0.5, legend=False)
    ax.set_title(var, fontweight="bold")
    ax.set_ylabel("%")
    ax.tick_params(axis="x", rotation=30, labelsize=8)
    ax.grid(axis="y", alpha=0.3)
    ax.spines[["top", "right"]].set_visible(False)

handles = [plt.Rectangle((0, 0), 1, 1, color=colors_g[g]) for g in cat_order]
fig.legend(handles, [cat_labels[g] for g in cat_order],
           loc="lower center", ncol=4, fontsize=9,
           bbox_to_anchor=(0.5, -0.03), framealpha=0.9)
plt.tight_layout()
plt.savefig("demografisch_profiel_situatiegroepen.png", dpi=150, bbox_inches="tight")
plt.show()

# Bewaar voor Excel export
_excel_demografie = profile_wide.copy()
print("\\nFiguur opgeslagen: demografisch_profiel_situatiegroepen.png")
"""

# ─────────────────────────────────────────────────────────────────────────────
# NEW CELL B – Variantiedecompositie levenssatisfactie
#   Welk % wordt verklaard door: domein-tevredenheid | toekomst | invloed?
# ─────────────────────────────────────────────────────────────────────────────
CELL_B = """\
import statsmodels.api as sm
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

# ════════════════════════════════════════════════════════════════════════════════
# ANALYSE B: Variantiedecompositie levenssatisfactie
# Hoeveel % verklaart: domein-tevredenheid (nu) | toekomstverwachting | invloed?
# ════════════════════════════════════════════════════════════════════════════════

weight_col = "Weegfactor_GRDR"

target_subj = [c for c in tevreden_df.columns if "leven in het algemeen" in c.lower()][0]
pred_subjects = [s for s in subjects if s != target_subj]

print(f"Doelvariabele: {target_subj}")
print(f"Aantal predictor-domeinen: {len(pred_subjects)}")

# ── Bouw regressie-dataframe ─────────────────────────────────────────────────
rdf = pd.DataFrame({"y": tevreden_df[target_subj]})

dom_cols = []
for s in pred_subjects:
    c = f"dom_{s[:50]}"
    rdf[c] = tevreden_df[s]
    dom_cols.append(c)

toe_cols = []
for s in pred_subjects:
    if s in toekomst_df.columns:
        c = f"toe_{s[:50]}"
        rdf[c] = toekomst_df[s]
        toe_cols.append(c)

inv_cols = list(invloed_dummy_cols)
for d in inv_cols:
    rdf[d] = df[d]

rdf = rdf.dropna(subset=["y"])
all_pred_cols = dom_cols + toe_cols + inv_cols
for c in all_pred_cols:
    rdf[c] = pd.to_numeric(rdf[c], errors="coerce")
    rdf[c] = rdf[c].fillna(rdf[c].mean())

y = rdf["y"].astype(float)
print(f"N respondenten (na dropna doelvariabele): {len(y)}")

def fit_r2(cols):
    X = sm.add_constant(rdf[cols].astype(float))
    return sm.OLS(y, X).fit().rsquared

r2_dom  = fit_r2(dom_cols)
r2_dt   = fit_r2(dom_cols + toe_cols)   if toe_cols else r2_dom
r2_dti  = fit_r2(dom_cols + toe_cols + inv_cols) if inv_cols else r2_dt

incr_dom = r2_dom
incr_toe = r2_dt  - r2_dom
incr_inv = r2_dti - r2_dt
unexpl   = 1.0    - r2_dti

print("\\n=== Hiërarchische R² ===")
print(f"  Model 1 – Domein-tevredenheid alleen : R² = {r2_dom:.3f}  ({r2_dom*100:.1f}%)")
print(f"  Model 2 – + Toekomstverwachting      : R² = {r2_dt:.3f}  (+{incr_toe*100:.1f}%)")
print(f"  Model 3 – + Ervaren invloed           : R² = {r2_dti:.3f}  (+{incr_inv*100:.1f}%)")
print(f"  Onverklaard                           :       {unexpl*100:.1f}%")

# ── Visualisatie ──────────────────────────────────────────────────────────────
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
fig.suptitle("Variantiedecompositie levenssatisfactie\\n"
             "Wat bepaalt tevredenheid met het leven in het algemeen?",
             fontsize=14, fontweight="bold")

comp_data = {
    "Domein-\\ntevredenheid (nu)": (incr_dom * 100, "#e74c3c"),
    "Toekomst-\\nverwachting":      (incr_toe * 100, "#3498db"),
    "Ervaren\\ninvloed":            (incr_inv * 100, "#9b59b6"),
    "Onverklaard":                  (unexpl   * 100, "#ecf0f1"),
}

left = 0.0
for label, (val, color) in comp_data.items():
    ax1.barh(["Levenssatisfactie"], [val], left=left, color=color,
             label=f"{label.replace(chr(10), ' ')} ({val:.1f}%)", edgecolor="white", linewidth=0.5)
    if val > 4:
        ax1.text(left + val / 2, 0, f"{val:.1f}%",
                 ha="center", va="center", fontsize=11, fontweight="bold",
                 color="white" if color != "#ecf0f1" else "#555")
    left += val
ax1.set_xlim(0, 100)
ax1.set_xlabel("Percentage totale variantie (%)")
ax1.set_title("Aandeel in verklaarde variantie")
ax1.legend(loc="upper right", fontsize=9, framealpha=0.9)
ax1.spines[["top", "right", "left"]].set_visible(False)
ax1.set_yticks([])
ax1.grid(axis="x", alpha=0.3)

# Incrementele R² als staafdiagram
labels_bar = ["Domein-\\ntevredenheid", "Toekomst-\\nverwachting", "Ervaren\\ninvloed"]
vals_bar   = [incr_dom * 100, incr_toe * 100, incr_inv * 100]
cols_bar   = ["#e74c3c", "#3498db", "#9b59b6"]
bars = ax2.bar(labels_bar, vals_bar, color=cols_bar, edgecolor="white", width=0.5)
for bar, val in zip(bars, vals_bar):
    ax2.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.4,
             f"{val:.1f}%", ha="center", va="bottom", fontsize=12, fontweight="bold")
ax2.set_ylabel("Uniek verklaarde variantie (%)")
ax2.set_title("Unieke bijdrage per predictor-groep")
ax2.tick_params(axis="x", labelsize=10)
ax2.grid(axis="y", alpha=0.3)
ax2.spines[["top", "right"]].set_visible(False)

plt.tight_layout()
plt.savefig("variantiedecompositie_levenssatisfactie.png", dpi=150, bbox_inches="tight")
plt.show()

# Bewaar voor Excel export
_excel_vardecomp = pd.DataFrame({
    "Predictor-groep": [
        "Domein-tevredenheid (nu)",
        "Toekomstverwachting (uniek extra)",
        "Ervaren invloed (uniek extra)",
        "Onverklaard",
    ],
    "R2_bijdrage_%": [round(v * 100, 2) for v in
                      [incr_dom, incr_toe, incr_inv, unexpl]],
    "Cumulatief_R2": [round(v, 4) for v in [r2_dom, r2_dt, r2_dti, 1.0]],
})
print("\\nFiguur opgeslagen: variantiedecompositie_levenssatisfactie.png")
print(_excel_vardecomp.to_string(index=False))
"""

# ─────────────────────────────────────────────────────────────────────────────
# NEW CELL C – Beschermende factoren
#   Wat beschermt tegen 'ontevreden_en_pessimistisch'?
# ─────────────────────────────────────────────────────────────────────────────
CELL_C = """\
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

# ════════════════════════════════════════════════════════════════════════════════
# ANALYSE C: Beschermende factoren
# Vergelijk tevreden_stabiel vs ontevreden_pessimistisch op domeinen + invloed
# ════════════════════════════════════════════════════════════════════════════════

situatie_col = "situatie_uw_leven_in_het_algemeen"
weight_col   = "Weegfactor_GRDR"
grp_a = "tevreden_en_stabiel"
grp_b = "ontevreden_en_pessimistisch"

def weighted_mean(series, weights):
    sub = pd.DataFrame({"v": series, "w": weights}).dropna()
    if len(sub) == 0:
        return np.nan
    return np.average(sub["v"].astype(float), weights=sub["w"].astype(float))

# ── Domein-tevredenheid: gemiddeld per groep ────────────────────────────────
domain_rows = []
for s in subjects:
    if "leven in het algemeen" in s.lower():
        continue
    if s not in tevreden_df.columns:
        continue
    for g, label in [(grp_a, "Tevreden & stabiel"), (grp_b, "Ontevreden & pessimistisch")]:
        mask_g = df[situatie_col] == g
        mu = weighted_mean(tevreden_df[s][mask_g], df.loc[mask_g, weight_col])
        domain_rows.append({"Domein": s, "Groep": label, "Gem_tevredenheid": mu})

domain_df = (
    pd.DataFrame(domain_rows)
    .pivot(index="Domein", columns="Groep", values="Gem_tevredenheid")
)
domain_df["Verschil (A-B)"] = (
    domain_df.get("Tevreden & stabiel", 0) - domain_df.get("Ontevreden & pessimistisch", 0)
)
domain_df = domain_df.sort_values("Verschil (A-B)", ascending=False)

# ── Invloed: % met (beetje of veel) invloed per groep ───────────────────────
invloed_rows = []
for d in invloed_dummy_cols:
    for g, label in [(grp_a, "Tevreden & stabiel"), (grp_b, "Ontevreden & pessimistisch")]:
        mask_g = df[situatie_col] == g
        mu = weighted_mean(df.loc[mask_g, d], df.loc[mask_g, weight_col])
        invloed_rows.append({"Invloed": d, "Groep": label, "Pct_met_invloed": (mu or 0) * 100})

invloed_df = (
    pd.DataFrame(invloed_rows)
    .pivot(index="Invloed", columns="Groep", values="Pct_met_invloed")
)
invloed_df["Verschil (A-B)"] = (
    invloed_df.get("Tevreden & stabiel", 0) - invloed_df.get("Ontevreden & pessimistisch", 0)
)
invloed_df = invloed_df.sort_values("Verschil (A-B)", ascending=False)

print("=== Gemiddelde domein-tevredenheid (gewogen) ===")
print(domain_df.round(2).to_string())
print("\\n=== % met invloed (gewogen) ===")
print(invloed_df.round(1).to_string())

# ── Visualisatie ──────────────────────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(20, max(8, len(domain_df) * 0.4 + 2)))
fig.suptitle(
    "Beschermende factoren: Tevreden & stabiel vs. Ontevreden & pessimistisch\\n"
    "(leven in het algemeen, gewogen)",
    fontsize=14, fontweight="bold",
)

# Links: vergelijking domein-scores als twee naast elkaar gelegen staven
ax = axes[0]
col_a = "Tevreden & stabiel"
col_b = "Ontevreden & pessimistisch"
y_pos = np.arange(len(domain_df))
bw = 0.38
ax.barh(y_pos + bw / 2, domain_df.get(col_a, pd.Series(dtype=float)).fillna(0),
        height=bw, color="#2ecc71", label=col_a, alpha=0.9)
ax.barh(y_pos - bw / 2, domain_df.get(col_b, pd.Series(dtype=float)).fillna(0),
        height=bw, color="#e74c3c", label=col_b, alpha=0.9)
ax.set_yticks(y_pos)
ax.set_yticklabels([s[:55] for s in domain_df.index], fontsize=9)
ax.set_xlabel("Gem. tevredenheid (1–10)")
ax.set_title("Domein-tevredenheid per groep", fontweight="bold")
ax.set_xlim(0, 10)
ax.axvline(7, color="gray", linewidth=0.8, linestyle=":")
ax.legend(fontsize=9, loc="lower right")
ax.grid(axis="x", alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)

# Rechts: verschilgrafiek (A – B), inclusief invloed
ax2 = axes[1]
dom_verschil = domain_df[["Verschil (A-B)"]].assign(Type="Domein-tevredenheid")
inv_verschil = invloed_df[["Verschil (A-B)"]].rename(
    index=lambda x: x.replace("_invloed", " (invloed)")).assign(Type="Ervaren invloed")
verschil_all = pd.concat([dom_verschil, inv_verschil]).sort_values("Verschil (A-B)")

bar_colors = verschil_all["Type"].map({
    "Domein-tevredenheid": "#e74c3c",
    "Ervaren invloed":     "#9b59b6",
})
ax2.barh(range(len(verschil_all)), verschil_all["Verschil (A-B)"],
         color=bar_colors, alpha=0.85, edgecolor="white")
ax2.set_yticks(range(len(verschil_all)))
ax2.set_yticklabels([s[:60] for s in verschil_all.index], fontsize=8)
ax2.axvline(0, color="black", linewidth=0.8, linestyle="--")
ax2.set_xlabel("Verschil tevreden_stabiel − ontevreden_pessimistisch\\n"
               "(groter = meer beschermend voor de stabiele groep)")
ax2.set_title("Beschermende factoren (gesorteerd)", fontweight="bold")
ax2.legend(handles=[
    Patch(color="#e74c3c", label="Domein-tevredenheid"),
    Patch(color="#9b59b6", label="Ervaren invloed"),
], fontsize=9, loc="lower right")
ax2.grid(axis="x", alpha=0.3)
ax2.spines[["top", "right"]].set_visible(False)

plt.tight_layout()
plt.savefig("beschermende_factoren.png", dpi=150, bbox_inches="tight")
plt.show()

# Bewaar voor Excel export
_excel_beschermend_dom    = domain_df.round(3).copy()
_excel_beschermend_invloed = invloed_df.round(2).copy()
print("\\nFiguur opgeslagen: beschermende_factoren.png")
"""

# ─────────────────────────────────────────────────────────────────────────────
# NEW CELL D – Excel export van ALLE analyses (voor de designer)
# ─────────────────────────────────────────────────────────────────────────────
CELL_D = """\
import pandas as pd
import numpy as np

# ════════════════════════════════════════════════════════════════════════════════
# EXPORT: alle analysesresultaten naar Excel (voor de designer)
# Elk figuur/analyse heeft een eigen tabblad.
# ════════════════════════════════════════════════════════════════════════════════

excel_path = "analyses_figuren_export.xlsx"
weight_col = "Weegfactor_GRDR"
cat_order  = ["tevreden_en_stabiel", "tevreden_en_pessimistisch",
              "ontevreden_en_optimistisch", "ontevreden_en_pessimistisch"]

with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:

    # ── 1. Situatieverdeling (stacked bar) ───────────────────────────────────
    records = []
    for col in situatie_cols:
        sub = df[[col, weight_col]].dropna(subset=[col, weight_col])
        wc  = sub.groupby(col)[weight_col].sum()
        pct = (wc / wc.sum() * 100).rename("Percentage_%")
        pct.index.name = "Situatiegroep"
        pct = pct.reset_index()
        pct.insert(0, "Domein", col.replace("situatie_", "").replace("_", " ").title())
        records.append(pct)
    pd.concat(records, ignore_index=True).to_excel(
        writer, sheet_name="01_Situatieverdeling", index=False)

    # ── 2. Correlatiematrix tevredenheid ─────────────────────────────────────
    corr_matrix_tevredenheid.round(3).to_excel(writer, sheet_name="02_Correlaties_tevredenheid")

    # ── 3. Cross-correlatie toekomst-tevredenheid ────────────────────────────
    corr_matrix_toekomst.round(3).to_excel(writer, sheet_name="03_Correlaties_toekomst")

    # ── 4. Demografisch profiel per situatiegroep (Analyse A) ─────────────────
    if "_excel_demografie" in dir():
        _excel_demografie.round(1).to_excel(writer, sheet_name="04_Demografisch_profiel")

    # ── 5. Variantiedecompositie (Analyse B) ─────────────────────────────────
    if "_excel_vardecomp" in dir():
        _excel_vardecomp.to_excel(writer, sheet_name="05_Variantie_decompositie", index=False)

    # ── 6. Beschermende factoren – domeinen (Analyse C) ──────────────────────
    if "_excel_beschermend_dom" in dir():
        _excel_beschermend_dom.to_excel(writer, sheet_name="06_Beschermende_domeinen")

    # ── 7. Beschermende factoren – invloed (Analyse C) ───────────────────────
    if "_excel_beschermend_invloed" in dir():
        _excel_beschermend_invloed.to_excel(writer, sheet_name="07_Beschermende_invloed")

    # ── 8. Random Forest feature importances (top 50) ─────────────────────────
    if "importances" in dir():
        imp_df = importances.reset_index()
        imp_df.columns = ["Feature", "Belang_RF"]
        imp_df.head(50).to_excel(writer, sheet_name="08_RF_factoren_belang", index=False)

    # ── 9. Gewogen % sociaal netwerk / bestaanszekerheid ────────────────────
    if "weighted_pcts" in dir():
        wpct_df = weighted_pcts.reset_index()
        wpct_df.columns = ["Variabele", "Gewogen_pct_ja"]
        wpct_df.to_excel(writer, sheet_name="09_Netwerk_bestaanszekerheid", index=False)

    # ── 10. Beschermende factoren: alle situatiegroepen vergeleken ────────────
    #   Voor elk domein: gewogen gemiddelde tevredenheid per situatiegroep
    dom_all_rows = []
    for s in subjects:
        if "leven in het algemeen" in s.lower():
            continue
        if s not in tevreden_df.columns:
            continue
        row = {"Domein": s}
        for g in cat_order:
            mask_g = df["situatie_uw_leven_in_het_algemeen"] == g
            sub = pd.DataFrame({
                "v": tevreden_df[s][mask_g],
                "w": df.loc[mask_g, weight_col],
            }).dropna()
            row[g] = (np.average(sub["v"].astype(float), weights=sub["w"].astype(float))
                      if len(sub) > 0 else np.nan)
        dom_all_rows.append(row)
    pd.DataFrame(dom_all_rows).round(2).to_excel(
        writer, sheet_name="10_Domeinscores_per_groep", index=False)

    # ── 11. Invloed per situatiegroep ─────────────────────────────────────────
    inv_all_rows = []
    for d in invloed_dummy_cols:
        row = {"Invloed_variabele": d}
        for g in cat_order:
            mask_g = df["situatie_uw_leven_in_het_algemeen"] == g
            sub = pd.DataFrame({
                "v": df.loc[mask_g, d].astype(float),
                "w": df.loc[mask_g, weight_col],
            }).dropna()
            row[g] = (np.average(sub["v"], weights=sub["w"]) * 100
                      if len(sub) > 0 else np.nan)
        inv_all_rows.append(row)
    pd.DataFrame(inv_all_rows).round(1).to_excel(
        writer, sheet_name="11_Invloed_per_groep", index=False)

    # ── 12. Groepsgroottes (gewogen N en %) ───────────────────────────────────
    size_rows = []
    for col in situatie_cols:
        sub = df[[col, weight_col]].dropna(subset=[col, weight_col])
        wc  = sub.groupby(col)[weight_col].sum()
        pct = wc / wc.sum() * 100
        for g in wc.index:
            size_rows.append({
                "Domein": col.replace("situatie_", "").replace("_", " ").title(),
                "Situatiegroep": g,
                "Gewogen_N": round(wc[g], 1),
                "Pct_%": round(pct[g], 1),
            })
    pd.DataFrame(size_rows).to_excel(writer, sheet_name="12_Groepsgroottes", index=False)

print(f"Excel export klaar: {excel_path}")
xf = pd.ExcelFile(excel_path)
print(f"Sheets ({len(xf.sheet_names)}):")
for s in xf.sheet_names:
    print(f"  - {s}")
"""

# ─────────────────────────────────────────────────────────────────────────────
# Voeg nieuwe cellen toe aan het notebook
# ─────────────────────────────────────────────────────────────────────────────
cells.append(new_md("## Analyse A – Demografisch profiel per situatiegroep"))
cells.append(new_cell(CELL_A))
cells.append(new_md("## Analyse B – Variantiedecompositie levenssatisfactie"))
cells.append(new_cell(CELL_B))
cells.append(new_md("## Analyse C – Beschermende factoren"))
cells.append(new_cell(CELL_C))
cells.append(new_md("## Export: alle figuurdata naar Excel"))
cells.append(new_cell(CELL_D))

# ─────────────────────────────────────────────────────────────────────────────
# Schrijf het gewijzigde notebook terug
# ─────────────────────────────────────────────────────────────────────────────
with open(NB_PATH, "w", encoding="utf-8") as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print("Done. analyses.ipynb is bijgewerkt.")
print(f"Totaal cellen: {len(cells)}")
