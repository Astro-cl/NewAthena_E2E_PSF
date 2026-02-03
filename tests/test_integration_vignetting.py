import tempfile
import os
import pandas as pd
import numpy as np

from main import load_gaussians_from_excel


def create_test_workbook(path):
    # MM_PSF sheet
    mm_psf = pd.DataFrame({
        'MM #': [1],
        'm_rad [arcsec]': [0.0],
        'm_azi [arcsec]': [0.0],
        'sigma_rad [arcsec]': [4.3],
        'sigma_azi [arcsec]': [4.3],
    })

    # A_eff with standard energy column name to enable vignetting
    aeff = pd.DataFrame({
        'MM #': [1],
        '0.25 keV': [1.0],
    })

    # MM configuration: MM 1 at x=1,y=0 (radial unit vector (1,0))
    mm_conf = pd.DataFrame({
        'MM #': [1],
        'Position #': [1],
        'x_MM [m]': [1.0],
        'y_MM [m]': [0.0],
        'r_MM [m]': [1.0],
    })

    # Gravity offload: provide rotx=1.0 arcsec at position 1
    grav = pd.DataFrame({
        'Position #': [1],
        'd_grav_rotx [arcsec]': [1.0],
        'd_grav_roty [arcsec]': [0.0],
    })

    # Vignetting sheets: layout A with first column delta and column '1' for position
    vig_azi = pd.DataFrame({
        'delta_arcsec': [-1.0, 0.0, 1.0],
        '1': [0.9, 1.0, 1.1],
    })
    vig_rad = pd.DataFrame({
        'delta_arcsec': [-1.0, 0.0, 1.0],
        '1': [0.95, 1.0, 1.1],
    })

    with pd.ExcelWriter(path, engine='openpyxl') as w:
        mm_psf.to_excel(w, sheet_name='MM_PSF', index=False)
        aeff.to_excel(w, sheet_name='A_eff', index=False)
        mm_conf.to_excel(w, sheet_name='MM configuration', index=False)
        grav.to_excel(w, sheet_name='Gravity offload', index=False)
        vig_azi.to_excel(w, sheet_name='Vignetting rotazi', index=False)
        vig_rad.to_excel(w, sheet_name='Vignetting rotrad', index=False)


def test_integration_vignetting_applies_weights():
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tmp.close()
    try:
        create_test_workbook(tmp.name)
        df = load_gaussians_from_excel(tmp.name)
        # After vignetting: base weight 1.0 * vig_rad(1.0)=1.1 * vig_azi(0.0)=1.0 => 1.1
        assert 'weight' in df.columns
        w = float(df.loc[0, 'weight'])
        assert np.isclose(w, 1.1, atol=1e-6)
    finally:
        try:
            os.remove(tmp.name)
        except Exception:
            pass
