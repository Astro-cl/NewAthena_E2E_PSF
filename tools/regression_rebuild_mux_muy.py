"""
Simple regression script to assert that `rebuild_df` computes non-zero
`mux`/`muy` when `MM #` keys and mm_config entries exist. This reproduces the
class of bug seen in combo 1769041290 where mismatched types caused zeroing.

Run with: python3 tools/regression_rebuild_mux_muy.py
"""
from pathlib import Path
import sys
import pandas as pd
import numpy as np

# Ensure workspace root is on sys.path for imports when run as a script
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from optimize_mm_rows import rebuild_df


def make_test_data():
    # Create params_df where MM # may be float-like (as produced by some reads)
    params = pd.DataFrame({
        "MM #": [1.0, 2.0, 3.0],
        "m_rad": [1e-6, 0.5e-6, 0.0],
        "m_azi": [0.0, 0.5e-6, 1e-6],
        "sigma_rad": [1e-8, 1e-8, 1e-8],
        "sigma_azi": [1e-8, 1e-8, 1e-8],
        "weight": [1.0, 1.0, 1.0],
    })

    # mm_config with integer MM # keys and non-zero x/y to make mux/muy non-zero
    mm_config = pd.DataFrame([
        {"MM #": 1, "x_MM [m]": 0.1, "y_MM [m]": 0.0, "r_MM [m]": 0.1},
        {"MM #": 2, "x_MM [m]": 0.0, "y_MM [m]": 0.1, "r_MM [m]": 0.1},
        {"MM #": 3, "x_MM [m]": -0.1, "y_MM [m]": 0.0, "r_MM [m]": 0.1},
    ])
    return params, mm_config


def run():
    params, mm_config = make_test_data()
    df = rebuild_df(params, mm_config)
    # Validate rebuilt mux/muy against direct polar->cartesian conversion
    print("rebuild_df output:\n", df[["MM #", "m_rad", "m_azi", "mux", "muy"]])
    # Build a quick mm_config_map (same convention as rebuild_df)
    mm_config_map = {
        int(r["MM #"]): {"x_MM": float(r["x_MM [m]"]), "y_MM": float(r["y_MM [m]"]), "r_MM": float(r["r_MM [m]"])}
        for _, r in mm_config.iterrows()
    }
    tol = 1e-15
    for _, row in df.iterrows():
        mm = int(row["MM #"]) if not pd.isna(row["MM #"]) else None
        if mm is None:
            continue
        cfg = mm_config_map.get(mm)
        if not cfg:
            continue
        x_mm = cfg["x_MM"]
        y_mm = cfg["y_MM"]
        r_mm = cfg["r_MM"] if cfg["r_MM"] != 0 else (abs(x_mm) if abs(x_mm) > 0 else 1e-9)
        u_rad_x = x_mm / r_mm
        u_rad_y = y_mm / r_mm
        u_azi_x = -y_mm / r_mm
        u_azi_y = x_mm / r_mm
        exp_mux = u_rad_x * float(row["m_rad"]) + u_azi_x * float(row["m_azi"])
        exp_muy = u_rad_y * float(row["m_rad"]) + u_azi_y * float(row["m_azi"])
        if not np.isclose(exp_mux, float(row["mux"]), atol=tol, rtol=0):
            raise SystemExit(f"Regression detected: mux mismatch for MM {mm}: expected {exp_mux} got {row['mux']}")
        if not np.isclose(exp_muy, float(row["muy"]), atol=tol, rtol=0):
            raise SystemExit(f"Regression detected: muy mismatch for MM {mm}: expected {exp_muy} got {row['muy']}")
    print("Regression test passed: rebuild_df mux/muy match expected polar->cartesian conversion.")


if __name__ == "__main__":
    run()
