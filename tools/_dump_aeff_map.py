import importlib.util
spec = importlib.util.spec_from_file_location('rs', 'tools/run_sensitivity.py')
rs = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rs)
mapd = rs.build_aeff_mapping()
print('col_names length', len(mapd.get('col_names',[])))
print('numeric_col_indices', mapd.get('numeric_col_indices'))
print('col_names samples:', mapd.get('col_names'))
