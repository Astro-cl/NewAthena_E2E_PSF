#!/usr/bin/env python3
"""Run the composite core+wing fit on a synthetic PSF.

Creates a two-component PSF (Gaussian core + pseudo-Voigt wing) and
calls `plot_sum` from `main.py` to perform the radial fits and save
diagnostic figures.
"""
import os
import numpy as np
import pandas as pd

import sys
import os as _os
# ensure project root is on sys.path so we can import main.py
_pr = _os.path.abspath(_os.path.join(_os.path.dirname(__file__), '..'))
if _pr not in sys.path:
    sys.path.insert(0, _pr)
import main

def main_run():
    arcsec_to_m = 12 * np.pi / 180 / 3600
    # Core ~0.6 arcsec FWHM, wing ~6 arcsec
    core_sigma_arcsec = 0.6
    wing_G_arcsec = 6.0

    sigc_m = core_sigma_arcsec * arcsec_to_m
    Gw_m = wing_G_arcsec * arcsec_to_m

    rows = [
        {
            'mux': 0.0,
            'muy': 0.0,
            'sigmax': sigc_m,
            'sigmay': sigc_m,
            'theta_degrees': 0.0,
            'distribution': 'gaussian',
            'weight': 0.7,
        },
        {
            'mux': 0.0,
            'muy': 0.0,
            'sigmax': Gw_m,
            'sigmay': Gw_m,
            'theta_degrees': 0.0,
            'distribution': 'pseudo-voigt',
            'alpha_azi': 0.2,
            'alpha_rad': 0.2,
            'weight': 0.3,
        },
    ]

    df = pd.DataFrame(rows)
    df.attrs['workbook_path'] = None

    os.makedirs('Figures', exist_ok=True)
    print('Running plot_sum on synthetic PSF (fast mode)...')
    # fast=True for quicker run; set fast=False for higher accuracy
    main.plot_sum(df, fast=True, normalize=True, debug=False)
    print('plot_sum finished. Check Figures/E2E_fit_composite.png')

if __name__ == '__main__':
    main_run()
