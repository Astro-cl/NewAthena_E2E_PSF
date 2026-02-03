import tempfile
import pandas as pd
import numpy as np
from main import load_gaussians_from_excel


def _make_workbook(path, rotz_arcsec):
    # MM configuration: one MM at x=1.0, y=0.0 (r=1)
    mm_conf = pd.DataFrame([
        {'MM #': 1, 'Position #': 1, 'x_MM [m]': 1.0, 'y_MM [m]': 0.0, 'z_MM [m]': 0.0, 'r_MM [m]': 1.0}
    ])

    # MM_PSF: minimal PSF row for MM #1
    mm_psf = pd.DataFrame([
        {
            'MM #': 1,
            'm_rad [arcsec]': 0.0,
            'm_azi [arcsec]': 0.0,
            'sigma_rad [arcsec]': 1.0,
            'sigma_azi [arcsec]': 1.0,
            'weight': 1.0,
        }
    ])

    # Alignment sheet: ensure rotz column is placed at index 6 (column G)
    # Build a header with 7 columns so column G holds 'd_align_rotz [arcsec]'
    align_cols = [
        'Position #',
        'd_align_rad [µm]',
        'd_align_azi [µm]',
        'd_align_z [µm]',
        'd_align_rotazi [arcsec]',
        'd_align_rotrad [arcsec]',
        'd_align_rotz [arcsec]',
    ]
    align_row = {c: 0 for c in align_cols}
    align_row['Position #'] = 1
    align_row['d_align_rotz [arcsec]'] = rotz_arcsec
    align_df = pd.DataFrame([align_row], columns=align_cols)

    with pd.ExcelWriter(path, engine='openpyxl') as w:
        mm_conf.to_excel(w, sheet_name='MM configuration', index=False)
        mm_psf.to_excel(w, sheet_name='MM_PSF', index=False)
        align_df.to_excel(w, sheet_name='Alignment', index=False)
        # Minimal A_eff sheet: two columns MM # and weight (column B used by loader)
        aeff = pd.DataFrame([
            {'MM #': 1, 'weight': 1.0}
        ])
        aeff.to_excel(w, sheet_name='A_eff', index=False)


def test_align_rotz_in_column_g_shifts_muy():
    tf_zero = tempfile.NamedTemporaryFile(prefix='test_align_zero_', suffix='.xlsx', delete=False)
    tf_non = tempfile.NamedTemporaryFile(prefix='test_align_20_', suffix='.xlsx', delete=False)
    tf_zero.close()
    tf_non.close()

    try:
        _make_workbook(tf_zero.name, rotz_arcsec=0.0)
        _make_workbook(tf_non.name, rotz_arcsec=20.0)

        df_zero = load_gaussians_from_excel(tf_zero.name, sheet='MM_PSF')
        df_non = load_gaussians_from_excel(tf_non.name, sheet='MM_PSF')

        assert 'muy' in df_zero.columns and 'muy' in df_non.columns

        muy0 = float(df_zero.loc[df_zero['MM #'] == 1, 'muy'].iloc[0])
        muyn = float(df_non.loc[df_non['MM #'] == 1, 'muy'].iloc[0])

        # Expect a positive increase in muy when rotz=+20 arcsec for an MM at x=1,y=0
        assert muyn > muy0

        # Expected shift ~ r * theta_rad where theta_rad = radians(20/3600)
        expected = np.radians(20.0 / 3600.0) * 1.0
        # Check within a small relative tolerance
        assert np.isclose(muyn - muy0, expected, rtol=1e-3, atol=1e-8)
    finally:
        import os
        try:
            os.unlink(tf_zero.name)
        except Exception:
            pass
        try:
            os.unlink(tf_non.name)
        except Exception:
            pass
