from pathlib import Path
import sys
sys.path.insert(0, 'sensitivity')
import run_batched as rb
from datetime import datetime
import pandas as pd

rb.INPUT_DIR.mkdir(parents=True, exist_ok=True)
rb.RESULTS_DIR.mkdir(parents=True, exist_ok=True)
rb.cleanup_input_dir()

baseline = rb.ROOT / 'Distributions' / 'Test_Distribution.xlsx'
std_mm_psf = rb.load_standard_mm_psf_defs(baseline)
std_alignment = rb.load_standard_alignment_defs(baseline)
std_thermal = rb.load_standard_thermal_defs(baseline)
std_gravity = rb.load_standard_gravity_defs(baseline)

# num_mm
try:
    mm_cfg = pd.read_excel(baseline, sheet_name='MM configuration', engine='openpyxl')
    num_mm = len(mm_cfg['MM #'].dropna()) if 'MM #' in mm_cfg.columns else 8
except Exception:
    num_mm = 8

sens_df = pd.read_excel(rb.SENS_DIR / 'sensitivity_input.xlsx', engine='openpyxl')
choices = {}
import re
for col in sens_df.columns:
    vals = []
    for v in sens_df[col].dropna():
        if isinstance(v, str) and (',' in v or '\n' in v):
            vals.extend([p.strip() for p in re.split('[,\n]', v) if p.strip()])
        else:
            vals.append(v)
    choices[col] = list(dict.fromkeys(vals))

from itertools import product
combos = [dict(zip(choices.keys(), prod)) for prod in product(*(choices[c] for c in choices))]
sel = combos[:3]
print('Generating', len(sel), 'workbooks...')
written = []
for i, combo in enumerate(sel, start=1):
    ts = datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')
    parts = [f"{k}={v}" for k, v in combo.items()]
    name_suffix = '_'.join(rb.sanitize_filename(p) for p in parts)
    filename = f"{ts}_{i}_{name_suffix}.xlsx"
    out_path = rb.INPUT_DIR / filename
    print(' -', out_path.name)
    rb.apply_combo_to_file(baseline, out_path, combo, std_mm_psf, std_alignment, std_thermal, std_gravity, num_mm)
    written.append(out_path)

# Inspect first workbook using main.load_gaussians_from_excel
sys.path.insert(0, '.')
import main
for p in written:
    print('\nInspecting', p.name)
    try:
        df = main.load_gaussians_from_excel(str(p))
        cols = [c for c in ['MM #','aeff_base','aeff_adjusted','aeff_vig_factor','weight'] if c in df.columns]
        print('Columns present:', cols)
        print(df[cols].head(8).to_string(index=False))
    except Exception as e:
        print('Load failed:', e)
