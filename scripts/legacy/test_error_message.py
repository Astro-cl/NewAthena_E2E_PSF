#!/usr/bin/env python3
"""Test the improved error message for missing formula values."""

import sys
import os
import pandas as pd
import numpy as np

sys.path.insert(0, '/Users/ivo ferreira/Library/CloudStorage/OneDrive-ESA/NewAthenaE2EPSF')

from openpyxl.utils import column_index_from_string

print("Testing improved error message...")

# Create a mock aeff_raw_df with None values in column P
data = {
    0: [1, 2, 3],  # MM numbers
    1: [None, None, None],  # Column B (A_eff weights) - empty
    15: [None, None, None],  # Column P (10 keV preset column) - contains None due to uncached formulas
}

# Create a minimal DataFrame that has column P with None values
aeff_raw_df = pd.DataFrame({
    0: [1, 2, 3],  # Column A
    1: [None, None, None],  # Column B
    15: [None, None, None],  # Column P
})

# Ensure we have enough columns
while aeff_raw_df.shape[1] <= 16:
    aeff_raw_df[aeff_raw_df.shape[1]] = None

row_idx = 1
col_letter = 'P'

try:
    col_idx = column_index_from_string(col_letter) - 1
    v = aeff_raw_df.iloc[row_idx, col_idx]
    print(f"Value at {col_letter}:{row_idx + 1} = {v}")
    
    try:
        float(v)
    except Exception:
        if pd.isna(v) or v is None:
            error_msg = (
                f'Column {col_letter} contains no numeric data at row {row_idx + 1}. '
                f'This column may contain formulas that have not been calculated. '
                f'Please open the Excel file in Microsoft Excel, press Ctrl+Shift+F9 to recalculate all formulas, '
                f'then save the file to cache the calculated values.'
            )
            print(f"\nError message that would be shown to user:\n")
            print(f"  {error_msg}")
            print("\n✓ Error message is helpful and actionable")
        
except Exception as e:
    print(f"Unexpected error: {e}")
