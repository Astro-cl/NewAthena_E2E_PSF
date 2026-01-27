#!/usr/bin/env python3
from pathlib import Path
import sys
sys.path.insert(0, str(Path('.').resolve()))
from main import load_gaussians_from_excel
from tools import run_sensitivity as rs
import numpy as np
import pandas as pd

BASE = rs.BASE_WORKBOOK
choice = "Fixed Sym Gaussian (sigma_rad=0.05 arcsec, sigma_azi=0.05 arcsec)"
print('Choice:', choice)

df = load_gaussians_from_excel(str(BASE), sheet='MM_PSF')
print('Original MM # sample:')
print(df[['MM #']].head().to_string())

ae = rs.build_aeff_mapping()
ae['mm_row_map'] = rs.load_mm_row_map(BASE)
standard_presets = rs.load_standard_mm_psf_presets(BASE)
rng = np.random.default_rng(12345)
df2 = rs.apply_mm_psf_choice_to_df(df.copy(), choice, ae, rng, standard_presets=standard_presets)

print('\nAfter apply_mm_psf_choice_to_df (meter-valued internals):')
cols = ['MM #','m_azi','m_rad','sigma_rad','sigma_azi','alpha_rad','alpha_azi']
for c in cols:
    if c not in df2.columns:
        df2[c] = None
print(df2[cols].head().to_string())

print('\nPrepared for main (arcsec):')
prepared = rs._prepare_df_for_main(df2)
cols2 = ['MM #','m_azi [arcsec]','m_rad [arcsec]','sigma_rad [arcsec]','sigma_azi [arcsec]']
for c in cols2:
    if c not in prepared.columns:
        prepared[c] = None
print(prepared[cols2].head().to_string())
