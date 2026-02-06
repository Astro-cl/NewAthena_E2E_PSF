#!/usr/bin/env python3
"""
Fix missing input workbooks for sensitivity run.

This script generates the missing Excel input workbooks for combos that 
only have CSV debug files. It uses the placed.xlsx files as templates
and adjusts the A_eff sheet for 1 keV combos.
"""

import pandas as pd
import re
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
INPUT_DIR = ROOT / 'sensitivity' / 'input'
RESULTS_FILE = ROOT / 'sensitivity' / 'results' / 'sensitivity_run_results.xlsx'

def parse_combo_filename(fname):
    """Parse a combo filename to extract parameters."""
    # Pattern: 20260204T200604Z_1_A_eff1_keV_MM_PSFFixed_Sym_Gaussian_0.01_Alignment0_Thermal0_Gravity_offload0.xlsx
    match = re.match(
        r'(\d+T\d+Z)_(\d+)_(.*)\.xlsx',
        fname.name
    )
    if match:
        ts, combo_id, rest = match.groups()
        return {'timestamp': ts, 'combo_id': int(combo_id), 'filename': fname, 'rest': rest}
    return None

def get_aeff_value_for_energy(aeff_sheet, energy):
    """Extract A_eff value for a given energy (1 keV, 7 keV, etc.)."""
    if aeff_sheet is None or not isinstance(aeff_sheet, pd.DataFrame):
        return None
    
    # Try to find the right column based on energy name
    for col in aeff_sheet.columns:
        if isinstance(col, str) and energy.lower() in col.lower():
            # Get mean value from this column
            vals = pd.to_numeric(aeff_sheet[col], errors='coerce')
            if vals.notna().any():
                return float(vals.mean())
    
    # Fallback: try to parse energy from column name as numeric
    for col in aeff_sheet.columns:
        if isinstance(col, str):
            # Look for pattern like "1 keV" or "7 keV"
            m = re.search(r'(\d+)\s*keV', col, re.IGNORECASE)
            if m:
                found_energy = int(m.group(1))
                target_energy = int(re.search(r'(\d+)', energy).group(1)) if re.search(r'(\d+)', energy) else None
                if found_energy == target_energy:
                    vals = pd.to_numeric(aeff_sheet[col], errors='coerce')
                    if vals.notna().any():
                        return float(vals.mean())
    
    # Default: return 1.0 if no specific value found
    return 1.0

def main():
    print("=" * 60)
    print("Fixing missing sensitivity input workbooks")
    print("=" * 60)
    
    # Load results to see which files are referenced
    try:
        results_df = pd.read_excel(RESULTS_FILE)
        print(f"Loaded {len(results_df)} results from {RESULTS_FILE}")
    except Exception as e:
        print(f"Failed to load results file: {e}")
        return
    
    # Find the 7 keV placed file to use as template
    placed_files = list(INPUT_DIR.glob('*_placed.xlsx'))
    if not placed_files:
        print("No placed.xlsx files found to use as template!")
        return
    
    print(f"Found {len(placed_files)} placed.xlsx files to use as templates")
    
    # Use the first placed file as template (7 keV one)
    template_file = placed_files[0]
    print(f"Using template: {template_file.name}")
    
    # Load template workbook structure
    try:
        template_wb = pd.ExcelFile(template_file)
        template_sheets = template_wb.sheet_names
        print(f"Template sheets: {template_sheets}")
    except Exception as e:
        print(f"Failed to load template: {e}")
        return
    
    # Get A_eff sheet from template
    aeff_sheet = None
    if 'A_eff' in template_sheets:
        aeff_sheet = pd.read_excel(template_file, sheet_name='A_eff')
        print(f"A_eff sheet columns: {list(aeff_sheet.columns)}")
    
    # Process each result row
    missing_files = []
    fixed_count = 0
    
    for idx, row in results_df.iterrows():
        input_file = row.get('input_file', '')
        combo_id = row.get('combo_id', idx + 1)
        error = row.get('error', '')
        a_eff = row.get('A_eff', '')
        
        # Check if this is an error due to missing file
        if not input_file:
            continue
            
        input_path = INPUT_DIR / input_file
        
        if not input_path.exists():
            missing_files.append({
                'combo_id': combo_id,
                'input_file': input_file,
                'a_eff': a_eff,
                'error': str(error)[:200]
            })
    
    print(f"\nFound {len(missing_files)} missing files")
    
    # Generate missing workbooks
    for item in missing_files:
        combo_id = item['combo_id']
        input_file = item['input_file']
        a_eff = item['a_eff']
        
        input_path = INPUT_DIR / input_file
        
        print(f"\n[{combo_id}] Creating: {input_file}")
        
        try:
            # Read template workbook
            with pd.ExcelWriter(input_path, engine='openpyxl') as writer:
                for sheet_name in template_sheets:
                    if sheet_name == 'A_eff':
                        # Modify A_eff sheet for this energy
                        if aeff_sheet is not None:
                            df = aeff_sheet.copy()
                            # Update values based on energy
                            energy_val = get_aeff_value_for_energy(aeff_sheet, a_eff)
                            if energy_val is not None and 'A_eff' in df.columns:
                                # Set all A_eff values to the appropriate value for this energy
                                df['A_eff'] = energy_val
                            df.to_excel(writer, sheet_name='A_eff', index=False)
                            print(f"  - Set A_eff values to {energy_val}")
                        else:
                            # Create default A_eff sheet
                            default_aeff = pd.DataFrame({
                                'MM #': list(range(1, 50)),
                                'A_eff': [1.0 if '1 keV' in str(a_eff) else 7.0] * 49
                            })
                            default_aeff.to_excel(writer, sheet_name='A_eff', index=False)
                            print(f"  - Created default A_eff sheet")
                    else:
                        # Copy other sheets from template
                        df = pd.read_excel(template_file, sheet_name=sheet_name)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"  - Copied sheet: {sheet_name}")
                
                fixed_count += 1
                print(f"  ✓ Created successfully")
                
        except Exception as e:
            print(f"  ✗ Failed: {e}")
    
    print("\n" + "=" * 60)
    print(f"Summary: Created {fixed_count} / {len(missing_files)} missing input files")
    print("=" * 60)
    
    # Also check for CSV files that need to be converted to proper input format
    csv_files = list(INPUT_DIR.glob('*_sampling_detailed_*.csv'))
    print(f"\nNote: {len(csv_files)} debug CSV sampling files exist")
    
    if fixed_count == 0 and len(missing_files) > 0:
        print("\nNo files were fixed. The missing files may be CSV references in results.")
        print("Checking if we need to update results to point to existing files...")
        
        # Check if _placed.xlsx files exist that could be used
        placed_by_combo = {}
        for pf in placed_files:
            # Extract combo info from filename
            name = pf.stem
            # Try to match to a combo
            for item in missing_files:
                if item['a_eff'] in name or name.endswith('_placed'):
                    combo_id = item['combo_id']
                    if combo_id not in placed_by_combo:
                        placed_by_combo[combo_id] = []
                    placed_by_combo[combo_id].append(pf)
        
        if placed_by_combo:
            print(f"\nFound {len(placed_by_combo)} placed.xlsx files that could be used as inputs")

if __name__ == '__main__':
    main()

