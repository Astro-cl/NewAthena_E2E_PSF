import os
import tempfile
import pandas as pd
import pytest

from main import load_gaussians_from_excel, load_aeff_weight_map_with_name


def make_test_workbook(path):
    with pd.ExcelWriter(path, engine='openpyxl') as w:
        # A_eff with a standard energy column
        aeff = pd.DataFrame({'MM #': [1, 2], '0.25 keV': [10.0, 20.0]})
        aeff.to_excel(w, sheet_name='A_eff', index=False)

        # MM_PSF minimal rows
        mmpsf = pd.DataFrame({
            'MM #': [1, 2],
            'm_rad [arcsec]': [0.0, 0.0],
            'm_azi [arcsec]': [0.0, 0.0],
            'sigma_rad [arcsec]': [0.01, 0.01],
            'sigma_azi [arcsec]': [0.01, 0.01],
            'distribution': ['gaussian', 'gaussian']
        })
        mmpsf.to_excel(w, sheet_name='MM_PSF', index=False)

        # MM configuration mapping to positions
        mmcfg = pd.DataFrame({'MM #': [1, 2], 'Position #': [1, 2], 'x_MM [m]': [1.0, 0.0], 'y_MM [m]': [0.0, 1.0], 'r_MM [m]': [1.0, 1.0]})
        mmcfg.to_excel(w, sheet_name='MM configuration', index=False)

        # perturbation sheets with zero rotation deltas (so totals = 0)
        # alignment no longer provides cartesian rot x/y; keep placeholders for other fields
        align = pd.DataFrame({'Position #': [1, 2], 'd_align_rad [µm]': [0.0, 0.0], 'd_align_azi [µm]': [0.0, 0.0]})
        align.to_excel(w, sheet_name='Alignment', index=False)
        grav = pd.DataFrame({'Position #': [1, 2], 'd_grav_rotx [arcsec]': [0.0, 0.0], 'd_grav_roty [arcsec]': [0.0, 0.0]})
        grav.to_excel(w, sheet_name='Gravity offload', index=False)
        therm = pd.DataFrame({'Position #': [1, 2], 'd_therm_rotx [arcsec]': [0.0, 0.0], 'd_therm_roty [arcsec]': [0.0, 0.0]})
        therm.to_excel(w, sheet_name='Thermal', index=False)

        # Vignetting sheets (Layout A): delta column then position columns '1','2'
        # Use polar vignetting sheets (rotazi/rotrad)
        vazi = pd.DataFrame({'delta_arcsec': [-1.0, 0.0, 1.0], '1': [0.9, 0.9, 0.9], '2': [0.8, 0.8, 0.8]})
        vazi.to_excel(w, sheet_name='Vignetting rotazi', index=False)
        vrad = pd.DataFrame({'delta_arcsec': [-1.0, 0.0, 1.0], '1': [0.95, 0.95, 0.95], '2': [0.85, 0.85, 0.85]})
        vrad.to_excel(w, sheet_name='Vignetting rotrad', index=False)


def test_aeff_and_vignetting_adjustment():
    tf = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tf.close()
    path = tf.name
    try:
        make_test_workbook(path)

        # Check load_aeff_weight_map returns mapping and detects column
        mapping, colname = load_aeff_weight_map_with_name(path)
        assert isinstance(mapping, dict)
        assert 1 in mapping and 2 in mapping
        assert colname is not None and '0.25' in str(colname)

        # Load via loader that applies vignetting; with zero rotx/roty totals we expect
        # weights to be multiplied by the vignetting value at delta=0: pos1 -> 0.9*0.95, pos2 -> 0.8*0.85
        df = load_gaussians_from_excel(path, sheet='MM_PSF')
        w = df.set_index('MM #')['weight'].to_dict()
        assert pytest.approx(10.0 * 0.9 * 0.95, rel=1e-6) == w[1]
        assert pytest.approx(20.0 * 0.8 * 0.85, rel=1e-6) == w[2]
    finally:
        try:
            os.remove(path)
        except Exception:
            pass
