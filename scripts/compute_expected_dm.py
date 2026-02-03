import pandas as pd
from pathlib import Path
import numpy as np

wb = Path('Distributions/NewTest_Distribution.xlsx')
xl = pd.ExcelFile(wb)
mm_to_pos = {}
mm_config_map = {}
if 'MM configuration' in xl.sheet_names:
    mmc = xl.parse('MM configuration', header=0)
    if 'MM #' in mmc.columns:
        for i, row in mmc.iterrows():
            mmn = row.get('MM #')
            if pd.isna(mmn):
                continue
            mmn_i = int(pd.to_numeric(mmn, errors='coerce'))
            pos = None
            if 'Position #' in mmc.columns and not pd.isna(row.get('Position #')):
                pos = int(pd.to_numeric(row.get('Position #'), errors='coerce'))
            if pos is None:
                pos = len(mm_to_pos) + 1
            mm_to_pos[mmn_i] = pos
            x = row.get('x_MM [m]') if 'x_MM [m]' in mmc.columns else row.get('x_MM') if 'x_MM' in mmc.columns else 0.0
            y = row.get('y_MM [m]') if 'y_MM [m]' in mmc.columns else row.get('y_MM') if 'y_MM' in mmc.columns else 0.0
            r = row.get('r_MM [m]') if 'r_MM [m]' in mmc.columns else row.get('r_MM') if 'r_MM' in mmc.columns else 0.0
            mm_config_map[mmn_i] = {'x_MM': float(pd.to_numeric(x, errors='coerce') or 0.0), 'y_MM': float(pd.to_numeric(y, errors='coerce') or 0.0), 'r_MM': float(pd.to_numeric(r, errors='coerce') or 0.0)}

print('Loaded', len(mm_to_pos), 'MM->pos entries')

# read Alignment
if 'Alignment' not in xl.sheet_names:
    print('No Alignment sheet')
    raise SystemExit(0)

af = xl.parse('Alignment', header=0)
# detect position column
pos_col = None
for c in af.columns:
    if 'position' in str(c).lower():
        pos_col = c
        break
if pos_col is None:
    print('No Position column found in Alignment')
    raise SystemExit(0)

# find d_align_z column
dz_col = None
for c in af.columns:
    if 'd_align_z' in str(c).lower():
        dz_col = c
        break
if dz_col is None:
    print('No d_align_z column found')
    raise SystemExit(0)

# build mapping pos->d_align_z (convert µm->m if >1)
pos_to_dz = {}
for _, row in af.iterrows():
    pos = row.get(pos_col)
    try:
        pos_i = int(pd.to_numeric(pos, errors='coerce'))
    except Exception:
        continue
    val = row.get(dz_col)
    num = pd.to_numeric(val, errors='coerce')
    if pd.isna(num):
        continue
    if abs(float(num)) > 1.0:
        # assume µm
        mz = float(num) * 1e-6
    else:
        mz = float(num)
    pos_to_dz[pos_i] = mz

print('\nSample pos->d_align_z (m):')
for p in sorted(list(pos_to_dz.keys()))[:20]:
    print(p, '->', pos_to_dz[p])

# compute dm for MM entries whose pos has alignment dz
print('\nMM #, Position, d_align_z[m], dm_x[m], dm_y[m] (sample first 40 MM entries whose pos has dz)')
count = 0
for mmn in sorted(mm_to_pos.keys()):
    pos = mm_to_pos[mmn]
    if pos not in pos_to_dz:
        continue
    dz = pos_to_dz[pos]
    cfg = mm_config_map.get(mmn, {'x_MM':0.0,'y_MM':0.0,'r_MM':0.0})
    x = cfg.get('x_MM',0.0)
    y = cfg.get('y_MM',0.0)
    r = cfg.get('r_MM', 0.0) or float(np.hypot(x,y))
    if r and r != 0.0:
        u_rad_x = float(x)/float(r)
        u_rad_y = float(y)/float(r)
    else:
        u_rad_x, u_rad_y = 1.0, 0.0
    dm_x = dz * u_rad_x
    dm_y = dz * u_rad_y
    print(mmn, ',', pos, ',', dz, ',', dm_x, ',', dm_y)
    count += 1
    if count >= 40:
        break

print('\nDone.')
