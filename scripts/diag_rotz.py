import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from main import load_gaussians_from_excel
import pandas as pd
import numpy as np

p='Distributions/NewTest_Distribution.xlsx'
print('Loading DF...')
df = load_gaussians_from_excel(p, sheet='MM_PSF')
print('DF columns:', list(df.columns))
# compute centroid (prefer adjusted A_eff when available)
if 'aeff_adjusted' in df.columns:
    w = df['aeff_adjusted'].to_numpy(dtype=float)
else:
    w = df['weight'].to_numpy(dtype=float)
wsum = w.sum() if w.size else 0.0
cx = (df['mux']*w).sum()/wsum if wsum else 0.0
cy = (df['muy']*w).sum()/wsum if wsum else 0.0
print(f'centroid mux,muy (m): {cx:.6e}, {cy:.6e}')
print('per-MM sample (first 12):')
for i,row in df.head(12).iterrows():
    mmn = int(row['MM #'])
    print(mmn, f"mux={row['mux']:.6e}", f"muy={row['muy']:.6e}", f"m_azi={row['m_azi']:.6e}")

# read alignment and mm config to compute expected m_azi shifts
mm = pd.read_excel(p, sheet_name='MM configuration', engine='openpyxl')
align = pd.read_excel(p, sheet_name='Alignment', engine='openpyxl')
# build align map
align_map={}
if 'Position #' in align.columns:
    tmp=align.copy(); tmp['Position #']=pd.to_numeric(tmp['Position #'], errors='coerce'); tmp=tmp[tmp['Position #'].notna()]
    for _,row in tmp.iterrows():
        pos=int(row['Position #'])
        d_align_rotz=float(row.get('d_align_rotz [arcsec]',0))
        if (d_align_rotz==0 or pd.isna(d_align_rotz)) and tmp.shape[1]>6:
            cand=row.iloc[6]
            try:
                if pd.notna(cand) and str(cand).strip()!='': d_align_rotz=float(cand)
            except Exception:
                pass
        align_map[pos]=d_align_rotz

print('Alignment sample:', {k:align_map[k] for k in list(align_map)[:5]})

print('\nExpected shifts per MM (r_MM * d_rotz_rad -> m_azi):')
for i,row in df.head(12).iterrows():
    mmn=int(row['MM #'])
    r_mm = float(mm.loc[mm['MM #']==mmn,'r_MM [m]'].iloc[0])
    pos = int(mm.loc[mm['MM #']==mmn,'Position #'].iloc[0])
    d_rotz = align_map.get(pos,0.0)
    d_rotz_rad = np.radians(d_rotz/3600.0)
    exp_m_azi = r_mm * d_rotz_rad
    print(mmn, f'r_mm={r_mm:.6e}', f'd_rotz={d_rotz}', f'exp_m_azi={exp_m_azi:.6e}')
