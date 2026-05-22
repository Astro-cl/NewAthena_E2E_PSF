import tempfile
import os
import pandas as pd
import numpy as np

from main import load_gaussians_from_excel


def create_workbook_with_alignment_preset(path):
    # MM_PSF minimal
    mm_psf = pd.DataFrame({
        'MM #': [1],
        'm_rad [arcsec]': [0.0],
        'm_azi [arcsec]': [0.0],
        'sigma_rad [arcsec]': [4.3],
        'sigma_azi [arcsec]': [4.3],
    })

    # A_eff standard energy
    aeff = pd.DataFrame({'MM #': [1], '0.25 keV': [1.0]})

    # MM configuration: MM 1 at x=1,y=0
    mm_conf = pd.DataFrame({
        'MM #': [1],
        'Position #': [1],
        'x_MM [m]': [1.0],
        'y_MM [m]': [0.0],
        'r_MM [m]': [1.0],
    })

    # Alignment sheet with many columns; place a non-empty value in column index 8
    # and place rotazi/rotrad in columns index 12 (M) and 13 (N).
    cols = []
    for i in range(15):
        cols.append(f'c{i}')
    align_row = [None] * 15
    align_row[0] = 1  # Position # at col 0
    align_row[8] = 'preset_selected'  # selection marker in I (index 8)
    # set rotazi at M (index 12) and rotrad at N (index 13)
    align_row[12] = 2.0  # d_align_rotazi (arcsec)
    align_row[13] = -1.0  # d_align_rotrad (arcsec)
    align_df = pd.DataFrame([align_row], columns=cols)
    # rename first column to 'Position #' to trigger Position# branch
    col_names = cols.copy()
    col_names[0] = 'Position #'
    align_df.columns = col_names

    # Gravity and Thermal zero cartesian rotations
    grav = pd.DataFrame({'Position #': [1], 'd_grav_rotx [arcsec]': [0.0], 'd_grav_roty [arcsec]': [0.0]})
    therm = pd.DataFrame({'Position #': [1], 'd_therm_rotx [arcsec]': [0.0], 'd_therm_roty [arcsec]': [0.0]})

    # Vignetting polar sheets: rotazi and rotrad
    # rotazi: at delta=2.0 -> 1.2
    vig_azi = pd.DataFrame({'delta_arcsec': [1.0, 2.0, 3.0], '1': [1.0, 1.2, 1.0]})
    # rotrad: at abs(delta)=1.0 -> 0.8 (table uses positive magnitudes)
    vig_rad = pd.DataFrame({'delta_arcsec': [0.0, 1.0, 2.0], '1': [1.0, 0.8, 0.9]})

    with pd.ExcelWriter(path, engine='openpyxl') as w:
        mm_psf.to_excel(w, sheet_name='MM_PSF', index=False)
        aeff.to_excel(w, sheet_name='A_eff', index=False)
        mm_conf.to_excel(w, sheet_name='MM configuration', index=False)
        align_df.to_excel(w, sheet_name='Alignment', index=False)
        grav.to_excel(w, sheet_name='Gravity offload', index=False)
        therm.to_excel(w, sheet_name='Thermal', index=False)
        vig_azi.to_excel(w, sheet_name='Vignetting rotazi', index=False)
        vig_rad.to_excel(w, sheet_name='Vignetting rotrad', index=False)


def test_alignment_preset_applies_polar_vignetting():
    tf = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tf.close()
    path = tf.name
    try:
        create_workbook_with_alignment_preset(path)
        df = load_gaussians_from_excel(path)
        assert 'weight' in df.columns
        w = float(df.loc[0, 'weight'])
        # expected multiplier = vig_azi(2.0)=1.2 * vig_rad(-1.0)=0.8 -> 0.96
        assert np.isclose(w, 0.96, atol=1e-6)
    finally:
        try:
            os.remove(path)
        except Exception:
            pass
