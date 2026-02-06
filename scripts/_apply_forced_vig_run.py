import sys
from pathlib import Path
import shutil
import pandas as pd
import numpy as np
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import main

repo = Path(__file__).resolve().parents[1]
input_dir = repo / 'sensitivity' / 'input'
files = sorted(list(input_dir.glob('*.xlsx')), key=lambda p: p.stat().st_mtime, reverse=True)
if not files:
    print('No generated workbooks found')
    sys.exit(1)
wb = files[0]
print('Using', wb)

# Copy to temp path
tmp = repo / ('forced_' + wb.name)
shutil.copy(wb, tmp)
print('Copied to', tmp)

# Load workbook sheets
xls = pd.read_excel(tmp, sheet_name=None, engine='openpyxl')
# Ensure perturbation sheets exist and inject small deltas per-position
npos = 100
# Alignment: ensure Position # column exists, create rows for first n positions
if 'Alignment' in xls:
    df_align = xls['Alignment'].copy()
    # If Position # not present, create simple positional table
    if 'Position #' not in df_align.columns:
        df_align = pd.DataFrame({'Position #': list(range(1, npos+1))})
    # Add rotazi/rotrad columns
    df_align['d_align_rotazi'] = df_align.get('d_align_rotazi', 0.0) + 0.3
    df_align['d_align_rotrad'] = df_align.get('d_align_rotrad', 0.0) + 0.1
else:
    df_align = pd.DataFrame({'Position #': list(range(1, npos+1)), 'd_align_rotazi': 0.3, 'd_align_rotrad': 0.1})
xls['Alignment'] = df_align

# Gravity offload: set small rotx/roty
if 'Gravity offload' in xls:
    df_grav = xls['Gravity offload'].copy()
    if 'Position #' not in df_grav.columns and 'MM #' not in df_grav.columns:
        df_grav = pd.DataFrame({'Position #': list(range(1, npos+1))})
    df_grav['d_grav_rotx'] = df_grav.get('d_grav_rotx', 0.0) + 0.5
    df_grav['d_grav_roty'] = df_grav.get('d_grav_roty', 0.0) + 0.2
else:
    df_grav = pd.DataFrame({'Position #': list(range(1, npos+1)), 'd_grav_rotx': 0.5, 'd_grav_roty': 0.2})
xls['Gravity offload'] = df_grav

# Thermal: leave zeros or small
if 'Thermal' in xls:
    df_therm = xls['Thermal'].copy()
    if 'Position #' not in df_therm.columns:
        df_therm = pd.DataFrame({'Position #': list(range(1, npos+1))})
    df_therm['d_therm_rotx'] = df_therm.get('d_therm_rotx', 0.0) + 0.0
    df_therm['d_therm_roty'] = df_therm.get('d_therm_roty', 0.0) + 0.0
else:
    df_therm = pd.DataFrame({'Position #': list(range(1, npos+1)), 'd_therm_rotx': 0.0, 'd_therm_roty': 0.0})
xls['Thermal'] = df_therm

# Write back modified workbook
with pd.ExcelWriter(tmp, engine='openpyxl') as writer:
    for sname, sdf in xls.items():
        try:
            sdf.to_excel(writer, sheet_name=sname, index=False)
        except Exception:
            # fallback: skip problematic sheet
            continue

# Now run main.load_gaussians_from_excel on tmp to apply vignetting and write B columns
try:
    df = main.load_gaussians_from_excel(str(tmp))
    print(df[['MM #','aeff_base','aeff_vig_factor','aeff_adjusted']].head(8))
except Exception as e:
    print('Error running loader:', e)

# Inspect Vignetting rotazi sheet B column
try:
    vdf = pd.read_excel(tmp, sheet_name='Vignetting rotazi', engine='openpyxl', header=None)
    print('\nVignetting rotazi first 12 rows (cols 0..2):')
    print(vdf.iloc[:12,0:3])
except Exception as e:
    print('Could not read Vignetting rotazi:', e)

try:
    vdf2 = pd.read_excel(tmp, sheet_name='Vignetting rotrad', engine='openpyxl', header=None)
    print('\nVignetting rotrad first 12 rows (cols 0..2):')
    print(vdf2.iloc[:12,0:3])
except Exception as e:
    print('Could not read Vignetting rotrad:', e)

print('\nModified workbook saved at', tmp)
