import pandas as pd
from pathlib import Path
import importlib.util
import re
spec = importlib.util.spec_from_file_location('rs', 'tools/run_sensitivity.py')
rs = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rs)

SENS_PATH = rs.SENS_PATH
sens = pd.read_excel(SENS_PATH, sheet_name=0, engine='openpyxl', header=0)
param_options = {}
for col in sens.columns:
    col_vals = []
    for cell in sens[col].tolist():
        if pd.isna(cell):
            continue
        s = str(cell).strip()
        if s == '':
            continue
        if s.lower() in {'-', 'nan', 'none'}:
            continue
        col_vals.append(s)
    if col_vals:
        param_options[col] = col_vals

print('Initial param_options:')
for k,v in param_options.items():
    print(k, len(v), v)

# emulate expansion
if 'A_eff' in param_options:
    vals = param_options['A_eff']
    expanded = []
    row_token_re = rs.re.compile(r'^(.*?)\s*\[row#\]\s*$', re.IGNORECASE)
    if len(vals) == 1:
        m0 = row_token_re.match(vals[0])
        if m0:
            cols = rs.build_aeff_mapping().get('col_names', [])
            cols = [c for c in cols if c and str(c).strip()]
            if cols:
                expanded.extend(cols)
            else:
                row_keys = sorted(rs.build_aeff_mapping().get('mm_row_map', {}).keys())
                base = m0.group(1).strip()
                for r in row_keys:
                    expanded.append(f"{base} [row{r}]")
    if not expanded:
        for v in vals:
            expanded.append(v)
    param_options['A_eff'] = expanded

print('\nExpanded A_eff:')
print('count', len(param_options['A_eff']))
print(param_options['A_eff'])

# compute combos
import itertools
keys = list(param_options.keys())
combos = list(itertools.product(*(param_options[k] for k in keys)))
print('\nTotal combos:', len(combos))
print('Example combos (first 10):')
for c in combos[:10]:
    print(dict(zip(keys,c)))
