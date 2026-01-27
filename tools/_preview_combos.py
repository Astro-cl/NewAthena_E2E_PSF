from pathlib import Path
import itertools, re, sys
sys.path.insert(0, str(Path('.').resolve()))
from tools import run_sensitivity as rs
SENS_PATH = rs.SENS_PATH
BASE_WORKBOOK = rs.BASE_WORKBOOK
import pandas as pd
sens = pd.read_excel(SENS_PATH, sheet_name=0, engine='openpyxl', header=0)
param_options = {}
for col in sens.columns:
    vals = sens[col].dropna().astype(str).map(str.strip)
    vals = [v for v in vals if v not in ['', '-', 'NaN', 'nan', 'None']]
    if vals:
        param_options[col] = vals
# expand A_eff
if 'A_eff' in param_options:
    vals = param_options['A_eff']
    expanded = []
    row_token_re = rs.re.compile(r'^(.*?)\s*\[row#\]\s*$', re.IGNORECASE)
    aeff_map_tmp = rs.build_aeff_mapping()
    mm_row_map = rs.load_mm_row_map(BASE_WORKBOOK)
    aeff_map_tmp['mm_row_map'] = mm_row_map
    numeric_idx = aeff_map_tmp.get('numeric_col_indices', [])
    total_mm_rows = len(mm_row_map) if mm_row_map else 0
    n_numeric = len(numeric_idx)
    max_row = n_numeric if n_numeric > 0 else total_mm_rows
    if max_row <= 0:
        max_row = total_mm_rows
    row_keys = [r for r in sorted(mm_row_map.keys()) if r <= max_row] if mm_row_map else []
    for v in vals:
        m = row_token_re.match(v)
        if m and row_keys:
            base = m.group(1).strip()
            for r in row_keys:
                expanded.append(f"{base} [row{r}]")
        else:
            expanded.append(v)
    param_options['A_eff'] = expanded
keys = list(param_options.keys())
combos = list(itertools.product(*(param_options[k] for k in keys)))
print('Expanded A_eff row keys:', row_keys)
print('Numeric A_eff columns indices:', numeric_idx)
print('Total combos:', len(combos))
print('Preview first 20 combos:')
for c in combos[:20]:
    print(dict(zip(keys, c)))
