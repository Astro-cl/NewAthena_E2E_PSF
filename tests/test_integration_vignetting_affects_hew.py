import tempfile
import os
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')

from main import load_gaussians_from_excel, plot_sum


def _write_workbook(path, mm_rows, aeff_map, vig_factors_by_pos):
    # mm_rows: list of dicts for MM_PSF rows
    # aeff_map: dict {mm: weight}
    # vig_factors_by_pos: dict pos -> factor (applies for both azimuthal and radial at delta=0)
    with pd.ExcelWriter(path, engine='openpyxl') as w:
        pd.DataFrame(mm_rows).to_excel(w, sheet_name='MM_PSF', index=False)

        # MM configuration: ensure MM -> Position mapping by row order
        mm_config = []
        for mm in sorted([r['MM #'] for r in mm_rows]):
            # place MM around unit circle for deterministic conversion
            if mm == 1:
                mm_config.append({'MM #': mm, 'x_MM [m]': 1.0, 'y_MM [m]': 0.0, 'r_MM [m]': 1.0})
            else:
                mm_config.append({'MM #': mm, 'x_MM [m]': 0.0, 'y_MM [m]': 1.0, 'r_MM [m]': 1.0})
        pd.DataFrame(mm_config).to_excel(w, sheet_name='MM configuration', index=False)

        # A_eff headerful sheet
        aeff_df = pd.DataFrame({'MM #': list(aeff_map.keys()), 'weight': list(aeff_map.values())})
        aeff_df.to_excel(w, sheet_name='A_eff', index=False)

        # Alignment: per-position rows with zero rotations (so interpolation at delta=0)
        align_rows = []
        for pos in sorted(vig_factors_by_pos.keys()):
            align_rows.append({'Position #': pos, 'd_align_rotazi [arcsec]': 0.0, 'd_align_rotrad [arcsec]': 0.0})
        pd.DataFrame(align_rows).to_excel(w, sheet_name='Alignment', index=False)

        # Vignetting sheets (per-position columns). Use delta values [-1, 0, 1]
        deltas = [-1.0, 0.0, 1.0]
        vig_df = {'delta': deltas}
        for pos, factor in vig_factors_by_pos.items():
            # constant factor across delta rows for simplicity
            vig_df[str(pos)] = [factor] * len(deltas)
        pd.DataFrame(vig_df).to_excel(w, sheet_name='Vignetting rotazi', index=False)
        pd.DataFrame(vig_df).to_excel(w, sheet_name='Vignetting rotrad', index=False)


def test_vignetting_changes_hew_by_removing_mm():
    # Two MMs: MM 1 at center, MM 2 offset by 10 arcsec in m_rad
    mm_rows = [
        {'MM #': 1, 'm_rad [arcsec]': 0.0, 'm_azi [arcsec]': 0.0, 'sigma_rad [arcsec]': 1.0, 'sigma_azi [arcsec]': 1.0},
        {'MM #': 2, 'm_rad [arcsec]': 10.0, 'm_azi [arcsec]': 0.0, 'sigma_rad [arcsec]': 1.0, 'sigma_azi [arcsec]': 1.0},
    ]

    with tempfile.TemporaryDirectory() as td:
        f_base = os.path.join(td, 'both_vig_1.xlsx')
        f_remove1 = os.path.join(td, 'vig_zero_pos1.xlsx')
        f_mm2_only = os.path.join(td, 'mm2_only.xlsx')

        # Case A: both MMs active, no vignetting (all factors 1)
        _write_workbook(f_base, mm_rows, {1: 1.0, 2: 1.0}, {1: 1.0, 2: 1.0})

        # Case B: apply vignetting that zeros-out MM1 (pos1 factor = 0)
        _write_workbook(f_remove1, mm_rows, {1: 1.0, 2: 1.0}, {1: 0.0, 2: 1.0})

        # Case C: simulate only MM2 by setting A_eff of MM1 to zero (no vignetting needed)
        _write_workbook(f_mm2_only, mm_rows, {1: 0.0, 2: 1.0}, {1: 1.0, 2: 1.0})

        # Load and compute metrics
        df_base = load_gaussians_from_excel(f_base)
        m_base = plot_sum(df_base, return_metrics_only=True, fast=True)

        df_remove1 = load_gaussians_from_excel(f_remove1)
        m_remove1 = plot_sum(df_remove1, return_metrics_only=True, fast=True)

        df_mm2 = load_gaussians_from_excel(f_mm2_only)
        m_mm2 = plot_sum(df_mm2, return_metrics_only=True, fast=True)

        # When MM1 is vignetted to zero, the HEW should match the single-MM2 scenario
        hew_remove = m_remove1.get('hew_origin_arcsec')
        hew_mm2 = m_mm2.get('hew_origin_arcsec')
        assert hew_remove is not None and hew_mm2 is not None

        # Allow small numerical tolerance due to grid/resolution differences
        assert np.isclose(hew_remove, hew_mm2, atol=1e-2)
