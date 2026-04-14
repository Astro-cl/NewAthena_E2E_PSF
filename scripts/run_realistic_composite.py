#!/usr/bin/env python3
"""Run composite fit on a realistic workbook under Distributions/.

Loads MM_PSF from the given workbook using `load_gaussians_from_excel`
and runs `plot_sum` in higher-accuracy mode to generate diagnostics.
"""
import os
import sys
_pr = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if _pr not in sys.path:
    sys.path.insert(0, _pr)

from main import load_gaussians_from_excel, plot_sum

def main_run(path):
    print(f"Loading workbook: {path}")
    df = load_gaussians_from_excel(path)
    # attach workbook_path so custom PSFs can be resolved
    df.attrs['workbook_path'] = path
    os.makedirs('Figures', exist_ok=True)
    # run with higher accuracy (fast=False)
    plot_sum(df, fast=False, normalize=True, debug=False)
    print('Done. Check Figures for diagnostic plots (E2E_fit*).')

if __name__ == '__main__':
    path = os.path.join('Distributions', 'reallistic.xlsx')
    if not os.path.exists(path):
        print('Workbook not found:', path)
        sys.exit(2)
    main_run(path)
