#!/usr/bin/env python3
import numpy as np
import sys
import os
# Ensure project root is on sys.path so local modules can be imported when running from tools/
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
from optimize_mm_rows import _load_base_params_from_workbook, hew_fast_approximate

file = sys.argv[1] if len(sys.argv)>1 else 'Distributions/Test_Just_MM_8.xlsx'
print('Loading', file)
base = _load_base_params_from_workbook(file)
print(base[['MM #','sigma_rad','sigma_azi']])

for _, row in base.iterrows():
    mm = int(row['MM #'])
    sigx = float(row['sigma_rad'])
    sigy = float(row['sigma_azi'])
    # Analytical for circular equal-sigma
    if abs(sigx - sigy) < 1e-12:
        sigma = sigx
        hew_analytic = 2.0 * sigma * np.sqrt(2.0 * np.log(2.0))
    else:
        # approximate effective sigma
        sigma = np.sqrt(0.5*(sigx**2 + sigy**2))
        hew_analytic = 2.0 * sigma * np.sqrt(2.0 * np.log(2.0))
    # Build single-MM df for hew_fast_approximate
    import pandas as pd
    df = pd.DataFrame([{
        'MM #': mm,
        'mux': 0.0,
        'muy': 0.0,
        'sigmax': sigx,
        'sigmay': sigy,
        'theta_degrees': 0.0,
        'weight': 1.0,
        'distribution': 'gaussian'
    }])
    hew_sim = hew_fast_approximate(df)
    print(f"MM {mm}: sigx={sigx:.3e} m sigy={sigy:.3e} m | analytic HEW={hew_analytic*1e6:.3f} um | sim HEW={hew_sim*1e6:.3f} um | ratio sim/analytic={hew_sim/hew_analytic:.6f}")
