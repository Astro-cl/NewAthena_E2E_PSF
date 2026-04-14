#!/usr/bin/env python3
"""Verify the cached A_eff values look reasonable."""

import sys
sys.path.insert(0, '/Users/ivo ferreira/Library/CloudStorage/OneDrive-ESA/NewAthenaE2EPSF')

import openpyxl as _openpyxl
import pandas as pd

print("Verifying cached A_eff values...")

# Load the file
wb = _openpyxl.load_workbook('Distributions/test10kev.xlsx', data_only=True)
ws = wb['A_eff']

data = []
for row in ws.iter_rows(values_only=True):
    data.append(row)

df = pd.DataFrame(data)

print("\nSample of filled A_eff preset columns:")
print("\nMM# | A_eff @0.2 keV | A_eff @1 keV | A_eff @10 keV | A_eff @12 keV")
print("-" * 70)

for row_idx in range(1, 16):  # First 15 MM rows
    mm_num = df.iloc[row_idx, 0]
    col_0_2kev = df.iloc[row_idx, 9]  # Column J (A_eff @0.2 keV)
    col_1kev = df.iloc[row_idx, 11]   # Column L (A_eff @1 keV)
    col_10kev = df.iloc[row_idx, 15]  # Column P (A_eff @10 keV)
    col_12kev = df.iloc[row_idx, 16]  # Column Q (A_eff @12 keV)
    
    print(f"{mm_num:3.0f} | {col_0_2kev:14} | {col_1kev:12} | {col_10kev:13} | {col_12kev:13}")

print("\n✓ Verification complete!")
print("\nNote: Values should be numeric (not None). Some may be 0 depending on the data.")
