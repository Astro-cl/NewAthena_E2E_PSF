#!/usr/bin/env python3
"""Check Row # distribution in MM configuration."""

import openpyxl

wb = openpyxl.load_workbook('Distributions/test10kev.xlsx')
mm_config = wb['MM configuration']

print("Sample of MM configuration (first 50 rows):")
print("\nPos# | Petal | Row# | MM#")
print("-" * 30)

row_dist = {}
for row_idx, row in enumerate(mm_config.iter_rows(min_row=2, max_row=52, min_col=1, max_col=4, values_only=True)):
    pos_num, petal, row_num, mm_num = row
    if mm_num is None:
        break
    print(f"{int(pos_num):3} | {int(petal):5} | {int(row_num):4} | {int(mm_num):3}")
    
    # Track distribution of Row # values
    row_key = int(row_num)
    if row_key not in row_dist:
        row_dist[row_key] = []
    row_dist[row_key].append(int(mm_num))

print("\n\nRow # distribution:")
for row_key in sorted(row_dist.keys()):
    print(f"  Row# {row_key}: {len(row_dist[row_key])} MMs - {row_dist[row_key][:5]}...")
