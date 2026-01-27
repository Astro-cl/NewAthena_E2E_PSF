#!/usr/bin/env python3
import pandas as pd, glob, os

res_path='sensivitiy/results/sensitivity_run_results.xlsx'
if os.path.exists(res_path):
    df=pd.read_excel(res_path)
    print('Results rows:', len(df))
    errs = df['error'].notna().sum() if 'error' in df.columns else 0
    print('Errors count:', errs)
    print('Sample rows:')
    for _,r in df.head(5).iterrows():
        vals = {k: (v if (not pd.isna(v)) else None) for k,v in r.items()}
        print(' -', vals)
else:
    print('Results file not found:', res_path)

# scan placed workbooks for non-numeric sigmas
bad_files=[]
files = glob.glob('sensivitiy/input/*_placed.xlsx') + glob.glob('sensivitiy/input/*.xlsx')
for f in files:
    try:
        mm=pd.read_excel(f, sheet_name='MM_PSF', engine='openpyxl')
    except Exception:
        continue
    cols=[c for c in mm.columns if isinstance(c,str) and 'sigma_rad' in c.lower()]
    cols2=[c for c in mm.columns if isinstance(c,str) and 'sigma_azi' in c.lower()]
    if not cols or not cols2:
        bad_files.append((f,'missing_sigma_cols'))
        continue
    sr=pd.to_numeric(mm.loc[:49,cols[0]], errors='coerce')
    sa=pd.to_numeric(mm.loc[:49,cols2[0]], errors='coerce')
    if sr.isnull().any() or sa.isnull().any() or (sr<=0).any() or (sa<=0).any():
        bad_files.append((f,'non_numeric_or_nonpositive'))

print('Files scanned:', len(files))
print('Files with potential template issues:', len(bad_files))
for bf,reason in bad_files:
    print('-', repr(bf), reason)
