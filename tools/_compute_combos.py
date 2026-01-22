import importlib.util, itertools, re
import pandas as pd
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

# expansion per current run_sensitivity logic
if 'A_eff' in param_options:
    vals = param_options['A_eff']
    expanded = []
    row_token_re = rs.re.compile(r'^(.*?)\s*\[row#\]\s*$', re.IGNORECASE)
    if len(vals) == 1:
        m0 = row_token_re.match(vals[0])
        if m0:
            cols_raw = rs.build_aeff_mapping().get('col_names', [])
            numeric_idx = rs.build_aeff_mapping().get('numeric_col_indices', [])
            cols = []
            for i in numeric_idx:
                name = cols_raw[i] if i < len(cols_raw) and str(cols_raw[i]).strip() else f'col{i}'
                cols.append(str(name).strip())
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

keys = list(param_options.keys())
combos = list(itertools.product(*(param_options[k] for k in keys)))
print('A_eff options:', param_options.get('A_eff'))
print('A_eff count:', len(param_options.get('A_eff',[])))
print('MM_PSF count:', len(param_options.get('MM_PSF',[])))
print('Total combos:', len(combos))
for c in combos[:10]:
    print(dict(zip(keys,c)))
