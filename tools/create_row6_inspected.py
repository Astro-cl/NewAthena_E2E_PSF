#!/usr/bin/env python3
from pathlib import Path
import pandas as pd
import os
import shutil
import importlib.util

# Load tools/run_sensitivity.py as module `rs` without requiring package import
spec = importlib.util.spec_from_file_location('rs', Path(__file__).resolve().parent / 'run_sensitivity.py')
rs = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rs)

ROOT = Path(__file__).resolve().parent.parent
BASE = ROOT / 'Distributions' / 'Test_Distribution.xlsx'
OUT_DIR = ROOT / 'Figures' / 'kept_workbooks'
OUT_DIR.mkdir(parents=True, exist_ok=True)

# read all sheets
wb = pd.read_excel(BASE, sheet_name=None, engine='openpyxl')
base_mm = wb.get('MM_PSF')
if base_mm is None:
    print('BASE workbook missing MM_PSF sheet')
    raise SystemExit(1)

# Determine MM # corresponding to row 6 in A_eff
mm_row_map = rs.load_mm_row_map(BASE)
mm_for_row6 = mm_row_map.get(6)
print('MM for row6 ->', mm_for_row6)

# Prepare modified MM_PSF: preserve m_rad and m_azi, set sigma_* to 0.05 arcsec and alphas to 0
out_mm = base_mm.copy()
# ensure MM # preserved
if 'MM #' not in out_mm.columns:
    raise SystemExit('MM_PSF missing MM #')
# set sigma columns
for col in ['sigma_rad [arcsec]', 'sigma_azi [arcsec]']:
    out_mm[col] = 0.05
# ensure alpha columns exist then set to zero
if 'alpha_rad' not in out_mm.columns:
    out_mm['alpha_rad'] = 0.0
else:
    out_mm['alpha_rad'] = 0.0
if 'alpha_azi' not in out_mm.columns:
    out_mm['alpha_azi'] = 0.0
else:
    out_mm['alpha_azi'] = 0.0

# Prepare Gravity offload sheet keyed by MM # with zeros
mm_vals = out_mm['MM #'].astype(int).tolist()
grav_rows = []
for m in mm_vals:
    grav_rows.append({'MM #': int(m), 'd_grav_x [µm]': 0.0, 'd_grav_y [µm]': 0.0, 'd_grav_z [µm]': 0.0, 'd_grav_rotz [arcsec]': 0.0})
grav_df = pd.DataFrame(grav_rows)

# Prepare A_eff: zero column B except for mm_for_row6
aeff = wb.get('A_eff')
if aeff is None:
    print('BASE workbook missing A_eff sheet')
    raise SystemExit(1)
# find second column name
cols = list(aeff.columns)
if len(cols) < 2:
    print('A_eff sheet has fewer than 2 columns; abort')
    raise SystemExit(1)
valcol = cols[1]
print('A_eff value column:', valcol)
orig_vals = aeff[valcol].copy()
# create new A_eff copy
out_aeff = aeff.copy()
out_aeff[valcol] = 0.0
if mm_for_row6 is not None:
    # mm_for_row6 may be a single int or a list of MM #s
    if isinstance(mm_for_row6, (list, tuple, set)):
        mask = out_aeff['MM #'].isin([int(x) for x in mm_for_row6])
    else:
        mask = out_aeff['MM #'] == int(mm_for_row6)
    if mask.any():
        out_aeff.loc[mask, valcol] = orig_vals[mask]
    else:
        print('Warning: MM for row6 not found in A_eff rows')

# Save to new workbook
import datetime, uuid
stamp = datetime.datetime.now().strftime('%Y%m%dT%H%M%S')
suffix = uuid.uuid4().hex[:8]
out_path = OUT_DIR / f'kept_custom_row6_{stamp}_{suffix}_Test_Distribution.xlsx'
with pd.ExcelWriter(out_path, engine='openpyxl') as w:
    # write all original sheets except those we overwrite
    for name, df in wb.items():
        if name == 'MM_PSF':
            out_mm.to_excel(w, sheet_name='MM_PSF', index=False)
        elif name == 'A_eff':
            out_aeff.to_excel(w, sheet_name='A_eff', index=False)
        else:
            try:
                df.to_excel(w, sheet_name=name, index=False)
            except Exception:
                # if non-tabular sheet, skip
                pass
    # write Gravity offload
    grav_df.to_excel(w, sheet_name='Gravity offload', index=False)

print('Wrote inspected workbook to', out_path)
print('You can open it to confirm MM_PSF sigma/alpha, Gravity offload presence, and A_eff weights.')
