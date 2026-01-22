import importlib.util
spec = importlib.util.spec_from_file_location('rs', 'tools/run_sensitivity.py')
rs = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rs)
import numpy as np
cols_raw = rs.build_aeff_mapping().get('col_names', [])
vals_raw = rs.build_aeff_mapping().get('col_values', [])
selected = []
for i, vals in enumerate(vals_raw):
    try:
        arr = np.asarray(vals, dtype=float)
    except Exception:
        continue
    if arr.size == 0:
        continue
    finite = np.isfinite(arr)
    if not finite.any():
        continue
    if np.sum(np.abs(arr[finite])) <= 0.0:
        continue
    name = cols_raw[i] if i < len(cols_raw) and str(cols_raw[i]).strip() else f'col{i}'
    selected.append((i, name))
print('selected count', len(selected))
for idx,name in selected:
    print(idx,repr(name))
