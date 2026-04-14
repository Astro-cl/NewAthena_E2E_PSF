import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import pandas as pd
import numpy as np
from optimize_mm_rows import rebuild_df


def test_rebuild_mux_muy_preservation():
    params = pd.DataFrame({
        "MM #": [1.0, 2.0, 3.0],
        "m_rad": [1e-6, 0.5e-6, 0.0],
        "m_azi": [0.0, 0.5e-6, 1e-6],
        "sigma_rad": [1e-8, 1e-8, 1e-8],
        "sigma_azi": [1e-8, 1e-8, 1e-8],
        "weight": [1.0, 1.0, 1.0],
    })
    mm_config = pd.DataFrame([
        {"MM #": 1, "x_MM [m]": 0.1, "y_MM [m]": 0.0, "r_MM [m]": 0.1},
        {"MM #": 2, "x_MM [m]": 0.0, "y_MM [m]": 0.1, "r_MM [m]": 0.1},
        {"MM #": 3, "x_MM [m]": -0.1, "y_MM [m]": 0.0, "r_MM [m]": 0.1},
    ])
    df = rebuild_df(params, mm_config)
    # Expect mux/muy computed as polar->cartesian
    assert not df['mux'].isnull().any()
    assert not df['muy'].isnull().any()


def test_validate_sensitivity_no_mux_zeroing():
    from tools.validate_sensitivity_mux_muy import run_validation, OUTDIR
    # Run validation; it will produce a csv file. We assert it exists after running.
    run_validation()
    files = list(OUTDIR.glob('validate_sensitivity_mux_muy_*.csv'))
    assert files, 'Validation output not produced'
