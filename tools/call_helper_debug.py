#!/usr/bin/env python3
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from tools import run_sensitivity as rs
from main import load_gaussians_from_excel
import numpy as np, json

try:
    BASE = rs.BASE_WORKBOOK
    df = load_gaussians_from_excel(str(BASE), sheet='MM_PSF')
    aeff = rs.build_aeff_mapping(); aeff['mm_row_map'] = rs.load_mm_row_map(BASE)
    std = rs.load_standard_mm_psf_presets(BASE)
    rng = np.random.default_rng(12345)
    choice = 'Fixed Sym Gaussian (sigma_rad=0.05 arcsec, sigma_azi=0.05 arcsec)'
    df2 = rs.apply_mm_psf_choice_to_df(df.copy(), choice, aeff, rng, standard_presets=std)
    combo = {'MM_PSF': choice, 'Alignment': '0', 'Gravity offload': '0', 'Thermal': '0'}
    print('Calling helper...')
    res = rs._run_plot_sum_subprocess(df2, mode='coarse', timeout_s=120.0, combo=combo)
    print('Result type:', type(res))
    print(json.dumps(res, indent=2))
except Exception as e:
    import traceback
    print('Exception in caller:', e)
    traceback.print_exc()
