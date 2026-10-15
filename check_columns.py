# -*- coding: utf-8 -*-
"""Check actual column names in the dataset."""
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
    df['_bronbestand'] = path.split('\\')[-1]
    return df

df = load_and_label(path)
df = df[df['Reactiestatus'] != 'Niet geantwoord'].copy()

print(f"Shape: {df.shape}")
print()

# All columns
with open('column_names.txt', 'w', encoding='utf-8') as f:
    for c in df.columns:
        f.write(c + '\n')
print("All columns written to column_names.txt")

# Key columns we care about
keywords = ['weeg', 'leeftijd', 'geslacht', 'opleiding', 'inkomen', 'koop', 'huur',
            'stad', 'platteland', 'gemeente', 'regio', 'krimp', 'dorps']
print("\n=== Relevant columns by keyword ===")
for kw in keywords:
    matches = [c for c in df.columns if kw.lower() in c.lower()]
    if matches:
        print(f"\n[{kw}]:")
        for m in matches[:5]:
            print(f"  {m!r}")

# Check situatie_ columns
print("\n[situatie_]:")
for c in df.columns:
    if c.startswith('situatie_'):
        print(f"  {c!r}")
