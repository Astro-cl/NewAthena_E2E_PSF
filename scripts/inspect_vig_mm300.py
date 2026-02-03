import pandas as pd
import numpy as np
wb='Distributions/NewTest_Distribution.xlsx'
print('WB:', wb)

# Read Alignment
try:
    df_align = pd.read_excel(wb, sheet_name='Alignment')
except Exception as e:
    print('Error reading Alignment:', e)
    raise
print('\nAlignment columns:', list(df_align.columns))
# find position column
pos_col=None
for c in df_align.columns:
    if str(c).lower().strip() in ('position','pos','position #','position#') or 'position' in str(c).lower():
        pos_col=c
        break
if pos_col is None:
    print('No explicit position column; showing head')
    print(df_align.head())
else:
    print('Position column detected:', pos_col)
    row = df_align[df_align[pos_col]==300]
    print('Rows for pos 300:', len(row))
    print(row.to_dict(orient='list'))
# find d_align_rotazi column
rotazi_col=None
for c in df_align.columns:
    if 'rotazi' in str(c).lower() or 'rot_azi' in str(c).lower() or 'd_align_rot' in str(c).lower():
        rotazi_col=c
        break
print('rotazi_col:', rotazi_col)
if rotazi_col is not None and pos_col is not None:
    val = df_align.loc[df_align[pos_col]==300, rotazi_col]
    print('d_align_rotazi at pos300 raw:', val.tolist())

# Read Vignetting rotazi
try:
    df_vig = pd.read_excel(wb, sheet_name='Vignetting rotazi')
except Exception as e:
    print('Error reading Vignetting rotazi:', e)
    raise
print('\nVignetting columns:', list(df_vig.columns))
# find first numeric column index
numeric_idx=None
for i,c in enumerate(df_vig.columns):
    if np.issubdtype(df_vig[c].dtype, np.number):
        numeric_idx=i
        break
print('first numeric column index:', numeric_idx)
if numeric_idx is None:
    print('No numeric delta column found; show head')
    print(df_vig.head(10))
else:
    xs = df_vig.iloc[:, numeric_idx].dropna().values
    print('delta xs sample:', xs[:10])
    # find column for pos 300
    pos_colname=None
    for c in df_vig.columns:
        if str(300) in str(c):
            pos_colname=c
            break
    if pos_colname is None:
        for c in df_vig.columns:
            if str(c).strip()==str(300):
                pos_colname=c
                break
    print('pos column name in vignetting sheet:', pos_colname)
    if pos_colname:
        ys = df_vig[pos_colname].values
        print('ys sample:', ys[:10])
        interp = np.interp(100, xs, ys, left=ys[0], right=ys[-1])
        print('interp at delta=100:', interp)
    else:
        print('Could not find pos column in vignetting sheet; showing head')
        print(df_vig.head())
