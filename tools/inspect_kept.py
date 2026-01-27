#!/usr/bin/env python3
import pandas as pd
from pathlib import Path
p = Path('Figures/kept_workbooks').glob('kept_*Test_Distribution*.xlsx')
files = sorted(list(p))
if not files:
    print('No kept_workbooks found')
    raise SystemExit(1)
fp = files[-1]
print('Inspecting', fp)
wb = pd.ExcelFile(fp, engine='openpyxl')
print('Sheets:', wb.sheet_names)

def show_sheet(name):
    try:
        df = pd.read_excel(fp, sheet_name=name, engine='openpyxl')
        print(f"--- Sheet: {name} shape={df.shape} cols={list(df.columns)}")
        print(df.head(12).to_string(index=False))
        return df
    except Exception as e:
        print(f"Could not read sheet {name}: {e}")
        return None

mm = show_sheet('MM_PSF')
grav = show_sheet('Gravity offload')
aeff = show_sheet('A_eff')

# Quick checks
if mm is not None:
    for col in ['m_rad [arcsec]','m_azi [arcsec]','sigma_rad [arcsec]','sigma_azi [arcsec]','alpha_rad','alpha_azi','MM #']:
        if col in mm.columns:
            vals = mm[col].head(12).tolist()
            print(f"{col} sample:", vals)

if aeff is not None:
    # show column B values (second column) per MM
    cols = list(aeff.columns)
    if len(cols) >= 2:
        print('A_eff first 12 rows (cols A,B):')
        print(aeff.iloc[:12, :2].to_string(index=False))
    else:
        print('A_eff sheet has <2 columns; raw head:')
        print(aeff.head(12).to_string(index=False))

print('\nDone')
