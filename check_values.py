# -*- coding: utf-8 -*-
"""Check actual data values for satisfaction and future columns."""
import pyreadstat, pandas as pd, numpy as np

path = r'O:\Projecten lopend\BW & kansengelijkheid\Data\Werkbestand BW en kansengelijkheid.sav'

def load_and_label(path):
    df, meta = pyreadstat.read_sav(path)
    if 'gemcode2021' in df.columns and 'gemcode2021' in meta.variable_value_labels:
        df['gemeente'] = df['gemcode2021'].map(meta.variable_value_labels['gemcode2021'])
    label_map = {c: l for c, l in zip(meta.column_names, meta.column_labels) if l}
    df.rename(columns=label_map, inplace=True)
    for col_original, col_renamed in label_map.items():
        if col_original == 'gemcode2021': continue
        if col_original in meta.variable_value_labels:
            value_map = meta.variable_value_labels[col_original]
            df[col_renamed] = df[col_renamed].map(value_map).fillna(df[col_renamed])
    dupes = df.columns[df.columns.duplicated()].unique()
    if len(dupes) > 0:
        cols = pd.Series(df.columns)
        for dup in cols[cols.duplicated()].unique():
            idxs = cols[cols == dup].index
            for i, idx in enumerate(idxs):
                if i > 0: cols[idx] = f'{dup}__{i+1}'
        df.columns = cols
    return df

df = load_and_label(path)
df = df[df['Reactiestatus'] != 'Niet geantwoord'].copy()

# Check tevredenheid column values
tev_cols = [c for c in df.columns if 'hoe tevreden u bent met' in c]
if tev_cols:
    c = tev_cols[0]
    print(f"Tevredenheid col: {c[:60]}")
    print(f"  dtype: {df[c].dtype}")
    print(f"  sample values: {df[c].dropna().head(10).tolist()}")
    print(f"  value_counts (first 5):")
    print(df[c].value_counts().head(5).to_string())

print()
# Check toekomst column values
toe_cols = [c for c in df.columns if 'komende 5 jaar' in c]
if toe_cols:
    c = toe_cols[0]
    print(f"Toekomst col: {c[:60]}")
    print(f"  dtype: {df[c].dtype}")
    print(f"  sample values: {df[c].dropna().head(10).tolist()}")
    print(f"  value_counts:")
    print(df[c].value_counts().head(5).to_string())

print()
# Check invloed column values
inv_cols = [c for c in df.columns if 'hoeveel invloed' in c.lower()]
if inv_cols:
    c = inv_cols[0]
    print(f"Invloed col: {c[:60]}")
    print(f"  dtype: {df[c].dtype}")
    print(f"  value_counts:")
    print(df[c].value_counts().head(5).to_string())
