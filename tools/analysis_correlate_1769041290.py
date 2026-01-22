import pandas as pd
from pathlib import Path
inpath = Path('Figures/debug_combo_1769041290_input_df.csv')
finpath = Path('Figures/debug_combo_1769041290_final_df.csv')
outpath = Path('Figures/analysis_1769041290_input_vs_final.csv')
if not inpath.exists() or not finpath.exists():
    print('Missing input/final debug CSVs; aborting')
    raise SystemExit(1)

di = pd.read_csv(inpath)
_df = pd.read_csv(finpath)

# choose key
def choose_key(d):
    for k in ['MM #','MM_or_index','MM','row','index']:
        if k in d.columns:
            return k
    return None

k1 = choose_key(di)
k2 = choose_key(_df)
if k1 and k1 == k2:
    key = k1
else:
    di = di.reset_index().rename(columns={'index':'__idx'})
    _df = _df.reset_index().rename(columns={'index':'__idx'})
    key = '__idx'

# find mux/muy columns
def find_col(d, pat):
    for c in d.columns:
        if pat.lower() in c.lower():
            return c
    return None

mux_in = find_col(di, 'mux')
muy_in = find_col(di, 'muy')
mux_fin = find_col(_df, 'mux')
muy_fin = find_col(_df, 'muy')

cols_in = [key]
if mux_in: cols_in.append(mux_in)
if muy_in: cols_in.append(muy_in)
cols_fin = [key]
if mux_fin: cols_fin.append(mux_fin)
if muy_fin: cols_fin.append(muy_fin)

mini_in = di[cols_in].copy()
mini_fin = _df[cols_fin].copy()

if mux_in: mini_in = mini_in.rename(columns={mux_in: 'mux_in'})
if muy_in: mini_in = mini_in.rename(columns={muy_in: 'muy_in'})
if mux_fin: mini_fin = mini_fin.rename(columns={mux_fin: 'mux_fin'})
if muy_fin: mini_fin = mini_fin.rename(columns={muy_fin: 'muy_fin'})

merged = mini_in.merge(mini_fin, on=key, how='outer')

for col in ['mux_in', 'mux_fin', 'muy_in', 'muy_fin']:
    if col in merged.columns:
        merged[col] = pd.to_numeric(merged[col], errors='coerce')

merged['mux_changed'] = False
merged['muy_changed'] = False
if 'mux_in' in merged.columns and 'mux_fin' in merged.columns:
    merged['mux_changed'] = (~merged['mux_in'].fillna(0).eq(merged['mux_fin'].fillna(0)))
if 'muy_in' in merged.columns and 'muy_fin' in merged.columns:
    merged['muy_changed'] = (~merged['muy_in'].fillna(0).eq(merged['muy_fin'].fillna(0)))

merged['mux_zeroed'] = False
merged['muy_zeroed'] = False
if 'mux_in' in merged.columns and 'mux_fin' in merged.columns:
    merged['mux_zeroed'] = (merged['mux_in'].abs() > 1e-15) & (merged['mux_fin'].eq(0.0))
if 'muy_in' in merged.columns and 'muy_fin' in merged.columns:
    merged['muy_zeroed'] = (merged['muy_in'].abs() > 1e-15) & (merged['muy_fin'].eq(0.0))

changed = merged[merged['mux_changed'] | merged['muy_changed']].copy()
changed.to_csv(outpath, index=False)

print('Wrote', outpath)
print('Total rows (input):', len(di))
print('Total rows (final):', len(_df))
print('Rows with mux change:', int(merged[merged['mux_changed']].shape[0]))
print('Rows with muy change:', int(merged[merged['muy_changed']].shape[0]))
print('Rows with mux zeroed (nonzero->0):', int(merged['mux_zeroed'].sum()))
print('Rows with muy zeroed (nonzero->0):', int(merged['muy_zeroed'].sum()))
