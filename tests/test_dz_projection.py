import pytest
import numpy as np
from main import compute_dm_from_dz


def test_dm_from_dz_with_r_mm():
    mm = {'x_MM': 0.5, 'y_MM': 0.5, 'r_MM': np.hypot(0.5, 0.5)}
    dm_x, dm_y = compute_dm_from_dz(mm, {}, 1e-6)
    # projection along radial unit vector (x and y equal) -> dm_x ~= dm_y
    assert pytest.approx(dm_x, rel=1e-6) == pytest.approx(dm_y, rel=1e-6)
    # magnitude should equal dz (unit radial vector projected)
    mag = float(np.hypot(dm_x, dm_y))
    assert mag == pytest.approx(1e-6, rel=1e-6)


def test_dm_from_dz_with_theta_only():
    mm = {'x_MM': 0.0, 'y_MM': 0.0, 'r_MM': 0.0}
    row = {'theta_degrees': 90.0}
    dm_x, dm_y = compute_dm_from_dz(mm, row, 2e-6)
    # theta=90 deg -> radial unit vector ~ (cos(90)=0, sin(90)=1)
    assert pytest.approx(dm_x, abs=1e-12) == 0.0
    assert pytest.approx(dm_y, rel=1e-6) == pytest.approx(2e-6, rel=1e-6)


def test_dm_from_dz_fallback():
    mm = {'x_MM': 0.0, 'y_MM': 0.0}
    dm_x, dm_y = compute_dm_from_dz(mm, {}, 5e-7)
    # fallback radial vector (1,0)
    assert pytest.approx(dm_x, rel=1e-6) == pytest.approx(5e-7, rel=1e-6)
    assert pytest.approx(dm_y, abs=1e-12) == 0.0
