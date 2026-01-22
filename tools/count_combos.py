#!/usr/bin/env python3
from pathlib import Path
import sys
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import pandas as pd
import re
from tools.run_sensitivity import SENS_PATH, BASE_WORKBOOK, build_aeff_mapping, load_mm_row_map

sens = pd.read_excel(SENS_PATH, sheet_name=0, engine='openpyxl', header=0)
param_options = {}
for col in sens.columns:
    vals = sens[col].dropna().astype(str).map(str.strip)
    vals = [v for v in vals if v not in ['', '-', 'NaN', 'nan', 'None']]
    if vals:
        param_options[col] = vals

# Expand A_eff [row#]
aeff_map = build_aeff_mapping()
aeff_map['mm_row_map'] = load_mm_row_map(BASE_WORKBOOK)
if 'A_eff' in param_options:
    vals = param_options['A_eff']
    expanded = []
    row_token_re = re.compile(r'^(.*?)\s*\[row#\]\s*$', re.IGNORECASE)
    if aeff_map.get('mm_row_map'):
        row_keys = sorted(aeff_map['mm_row_map'].keys())
    else:
        row_keys = []
    for v in vals:
        m = row_token_re.match(v)
        if m and row_keys:
            base = m.group(1).strip()
            for r in row_keys:
                expanded.append(base + ' [row' + str(r) + ']')
        else:
            expanded.append(v)
    param_options['A_eff'] = expanded

counts = {k: len(v) for k, v in param_options.items()}
# compute total combos
from math import prod
if counts:
    total = prod(counts.values())
else:
    total = 0

print('Parameter counts:')
for k in sorted(counts.keys()):
    print(' - {}: {}'.format(k, counts[k]))
print('\nTotal combinations (expanded):', total)
