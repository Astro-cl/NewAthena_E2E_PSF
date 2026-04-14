import tempfile
from pathlib import Path
import pandas as pd

from sensitivity.sensitivity_run import write_multisheet_csv
from main import parse_multisheet_csv


def test_multisheet_csv_roundtrip_preserves_aeff_adjusted():
    # Prepare sample A_eff and MM_PSF sheets
    aeff = pd.DataFrame({
        'MM #': [1, 2, 3],
        'A_eff': [10.0, 20.0, 30.0],
        'aeff_adjusted': [9.0, 18.0, 27.0],
        'aeff_vig_factor': [0.9, 0.9, 0.9],
    })

    mm = pd.DataFrame({
        'MM #': [1, 2, 3],
        'sigma_rad [arcsec]': [1.0, 2.0, 3.0],
        'aeff_base': [10.0, 20.0, 30.0],
        'aeff_adjusted': [9.0, 18.0, 27.0],
        'aeff_vig_factor': [0.9, 0.9, 0.9],
    })

    sheets = {'A_eff': aeff, 'MM_PSF': mm}

    with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as tf:
        out_path = Path(tf.name)

    try:
        write_multisheet_csv(sheets, out_path)
        parsed = parse_multisheet_csv(str(out_path))

        # A_eff should be present and should NOT contain the aeff_adjusted column (writer sanitizes it)
        assert 'A_eff' in parsed
        aef = parsed['A_eff']
        assert 'aeff_adjusted' not in [c.lower() for c in aef.columns]

        # MM_PSF should contain weight equal to the aeff_adjusted values
        assert 'MM_PSF' in parsed
        mm_r = parsed['MM_PSF']
        # normalize numeric comparison
        weight = mm_r['weight'].astype(float).to_list()
        expected = [9.0, 18.0, 27.0]
        assert all(abs(w - e) < 1e-8 for w, e in zip(weight, expected))
    finally:
        try:
            out_path.unlink()
        except Exception:
            pass
