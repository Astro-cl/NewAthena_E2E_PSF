import openpyxl as _openpyxl
import pandas as pd

# Mimic the fixed loading code
try:
    path = 'Distributions/test10kev.xlsx'
    
    # Load A_eff sheet using openpyxl directly to get cached formula values
    wb = _openpyxl.load_workbook(path, data_only=True)
    ws = wb['A_eff']
    
    # Convert worksheet to DataFrame
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append(row)
    aeff_raw = pd.DataFrame(data)
    
    print(f"✓ Successfully loaded A_eff sheet")
    print(f"  Shape: {aeff_raw.shape}")
    print(f"  Sample headers (row 0): {aeff_raw.iloc[0, :5].tolist()}")
    
    # Try to load presets (like the GUI does)
    aeff_standard_presets = {}
    
    # Find preset name and value columns
    name_col = None
    values_col = None
    
    for r in range(min(3, aeff_raw.shape[0])):
        for c in range(aeff_raw.shape[1]):
            cell = aeff_raw.iloc[r, c]
            try:
                s = str(cell).strip().lower()
            except Exception:
                s = ''
            if not s:
                continue
            if 'standard' in s and name_col is None:
                name_col = c
            if 'value' in s and values_col is None:
                values_col = c
    
    if name_col is None:
        name_col = 3
    if values_col is None:
        values_col = 4
    
    print(f"\n✓ Found preset columns: name_col={name_col}, values_col={values_col}")
    print(f"  Column {name_col}: {aeff_raw.iloc[0, name_col]}")
    print(f"  Column {values_col}: {aeff_raw.iloc[0, values_col]}")
    
    # Load presets
    row_idx = 1
    preset_count = 0
    while row_idx < aeff_raw.shape[0]:
        name = aeff_raw.iloc[row_idx, name_col] if name_col < aeff_raw.shape[1] else None
        if pd.isna(name) or str(name).strip() == '':
            break
        expr = aeff_raw.iloc[row_idx, values_col] if values_col < aeff_raw.shape[1] else None
        if pd.isna(expr) or str(expr).strip() == '':
            break
        
        preset_name = str(name).strip()
        preset_expr = str(expr).strip()
        aeff_standard_presets[preset_name] = preset_expr
        preset_count += 1
        
        if preset_count <= 10:
            print(f"  Preset: '{preset_name}' -> '{preset_expr}'")
        row_idx += 1
    
    print(f"\n✓ Loaded {preset_count} presets")
    print("\nFirst 10 presets:")
    for i, (name, expr) in enumerate(list(aeff_standard_presets.items())[:10]):
        print(f"  {name:30} -> {expr}")
        
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()
