import importlib.util, re
spec = importlib.util.spec_from_file_location('rs', 'tools/run_sensitivity.py')
rs = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rs)
mapd = rs.build_aeff_mapping()
vals=['1 keV [row#]']
row_token_re = re.compile(r'^(.*?)\s*\[row#\]\s*$', re.IGNORECASE)
m0 = row_token_re.match(vals[0])
print('m0',bool(m0))
cols_raw = mapd.get('col_names', [])
numeric_idx = mapd.get('numeric_col_indices', [])
cols=[]
for i in numeric_idx:
    name = cols_raw[i] if i < len(cols_raw) and str(cols_raw[i]).strip() else f'col{i}'
    cols.append(str(name).strip())
print('numeric_idx',numeric_idx)
print('cols len',len(cols))
print(cols)
