#!/usr/bin/env python3
"""Test that the A_eff sheet loading and preset evaluation works correctly."""

import sys
import os
sys.path.insert(0, '/Users/ivo ferreira/Library/CloudStorage/OneDrive-ESA/NewAthenaE2EPSF')

import openpyxl as _openpyxl
import pandas as pd
import re
import ast
import numpy as np
from openpyxl.utils import column_index_from_string

# Simulate what the fixed GUI code does
print("=" * 70)
print("TESTING: A_eff Sheet Loading and Preset Evaluation")
print("=" * 70)

try:
    path = 'Distributions/test10kev.xlsx'
    
    # Step 1: Load A_eff sheet using openpyxl with data_only=True
    print("\n1. Loading A_eff sheet...")
    wb = _openpyxl.load_workbook(path, data_only=True)
    ws = wb['A_eff']
    
    # Convert worksheet to DataFrame
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append(row)
    aeff_raw = pd.DataFrame(data)
    
    print(f"   ✓ Loaded successfully, shape: {aeff_raw.shape}")
    
    # Step 2: Load standard presets
    print("\n2. Loading standard presets...")
    aeff_standard_presets = {}
    
    # Find preset name and value columns
    name_col = 3  # Column D (0-indexed)
    values_col = 4  # Column E (0-indexed)
    
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
        row_idx += 1
    
    print(f"   ✓ Loaded {preset_count} presets")
    
    # Step 3: Test evaluating a simple preset (column reference)
    print("\n3. Testing preset evaluation for '10 keV' preset...")
    preset_name = '10 keV'
    if preset_name in aeff_standard_presets:
        preset_expr = aeff_standard_presets[preset_name]
        print(f"   Preset '{preset_name}' expression: '{preset_expr}'")
        
        # Looking for MM #1 data
        mm_number = 1
        row_idx_for_mm = None
        for r in range(1, aeff_raw.shape[0]):
            mm = aeff_raw.iloc[r, 0]
            if pd.notna(mm):
                try:
                    if int(float(mm)) == mm_number:
                        row_idx_for_mm = r
                        break
                except Exception:
                    pass
        
        if row_idx_for_mm is not None:
            print(f"   Found MM #{mm_number} at row index {row_idx_for_mm}")
            
            # Parse the preset expression (should be 'P' for 10 keV)
            s = str(preset_expr).strip()
            m = re.fullmatch(r'([A-Za-z]+)', s.replace(' ', ''))
            if m:
                col_letter = m.group(1).upper()
                print(f"   Column letter from expression: '{col_letter}'")
                
                # Get the value from the column
                col_idx = column_index_from_string(col_letter) - 1
                v = aeff_raw.iloc[row_idx_for_mm, col_idx]
                
                print(f"   Value at column {col_letter}, row {row_idx_for_mm + 1}: {v} (type: {type(v).__name__})")
                
                if pd.isna(v) or v is None:
                    print(f"   ⚠️  WARNING: Value is NaN/None - formulas may not have been cached in Excel!")
                    print(f"      The Excel file needs to be opened in Excel and saved to cache formula results.")
                else:
                    print(f"   ✓ Value is numeric: {float(v)}")
        else:
            print(f"   ✗ Could not find MM #{mm_number} in sheet")
    else:
        print(f"   ✗ Preset '{preset_name}' not found")
    
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"✓ A_eff sheet loading: SUCCESS")
    print(f"✓ Preset loading: SUCCESS ({preset_count} presets found)")
    print(f"✓ Preset evaluation: {'check output above for column P value'}")
    print("\nNote: If column P value is None/NaN, the Excel file needs to be")
    print("opened and saved in Microsoft Excel to cache the formula results.")
    
except Exception as e:
    print(f"\n✗ ERROR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
