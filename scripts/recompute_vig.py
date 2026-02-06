import pandas as pd
import numpy as np
xls='Distributions/TestDistribution6_working.xlsx'
vdf = pd.read_excel(xls, sheet_name='Vignetting rotazi', engine='openpyxl', header=None)
# build ys_by_pos_azi like main
ys_by_pos_azi = {}
if vdf is not None and not vdf.empty and vdf.shape[1] >= 11:
    for _, r in vdf.iterrows():
        try:
            cfg_row = r.iloc[7]
            if pd.isna(cfg_row):
                continue
            cfg_row = int(float(cfg_row))
        except Exception:
            continue
        energy_marker = r.iloc[9] if vdf.shape[1] > 9 else None
        try:
            xval = float(r.iloc[8]) * 60.0 if not pd.isna(r.iloc[8]) else None
        except Exception:
            xval = None
        try:
            yval = float(r.iloc[10]) if vdf.shape[1] > 10 and not pd.isna(r.iloc[10]) else None
        except Exception:
            yval = None
        if xval is None or yval is None:
            continue
        key_str = (cfg_row, str(energy_marker).strip())
        key_num = None
        try:
            key_num = (cfg_row, float(energy_marker))
        except Exception:
            key_num = None
        if key_str not in ys_by_pos_azi:
            ys_by_pos_azi[key_str] = {'xs': [], 'ys': []}
        ys_by_pos_azi[key_str]['xs'].append(xval)
        ys_by_pos_azi[key_str]['ys'].append(yval)
        if key_num is not None:
            if key_num not in ys_by_pos_azi:
                ys_by_pos_azi[key_num] = {'xs': [], 'ys': []}
            ys_by_pos_azi[key_num]['xs'].append(xval)
            ys_by_pos_azi[key_num]['ys'].append(yval)
# sort
for k, v in list(ys_by_pos_azi.items()):
    order = np.argsort(v['xs'])
    xs_sorted = np.array(v['xs'], dtype=float)[order]
    ys_sorted = np.array(v['ys'], dtype=float)[order]
    ys_by_pos_azi[k] = (xs_sorted, ys_sorted)

# find cfg_row for pos 1
mmc = pd.read_excel(xls, sheet_name='MM configuration', engine='openpyxl')
if 'Position #' in mmc.columns:
    rows = mmc.index[mmc['Position #']==1].tolist()
    cfg_row = rows[0]+1 if rows else 1
else:
    cfg_row = 1

sel_energy = None
# check vdf cell C2
try:
    cand = vdf.iat[1,2]
    if cand is not None:
        try:
            sel_energy = float(str(cand).strip())
        except Exception:
            import re
            m = re.search(r"(\d+(?:\.\d*)?)", str(cand))
            if m:
                sel_energy = float(m.group(1))
except Exception:
    pass

print('cfg_row', cfg_row, 'sel_energy', sel_energy)
keyn = (cfg_row, float(sel_energy))
print('keyn in ys_by_pos_azi:', keyn in ys_by_pos_azi)
if keyn in ys_by_pos_azi:
    xs_use, ys_use = ys_by_pos_azi[keyn]
    print('xs sample (arcsec):', xs_use[:10])
    print('ys sample:', ys_use[:10])
    # suppose rot_azi is 60 arcsec
    print('interp at 60:', float(np.interp(60.0, xs_use, ys_use)))
    print('interp at 120:', float(np.interp(120.0, xs_use, ys_use)))
else:
    # list keys
    print('available keys sample:', list(ys_by_pos_azi.keys())[:10])
