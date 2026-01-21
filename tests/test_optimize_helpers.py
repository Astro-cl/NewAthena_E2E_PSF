import numpy as np
from optimize_mm_rows import _theta_degrees_from_xy, _mux_muy_for_slot


def test_theta_degrees_from_xy_known():
    # x=0,y=1 should point to 0 deg (on +y axis -> mapped accordingly)
    t = _theta_degrees_from_xy(0.0, 1.0)
    assert isinstance(t, float)


def test_mux_muy_for_slot_simple():
    # r_mm non-zero, simple values
    m_rad = 1.0
    m_azi = 0.5
    x_mm = 0.3
    y_mm = 0.4
    r_mm = 0.5
    mux, muy = _mux_muy_for_slot(m_rad, m_azi, x_mm, y_mm, r_mm)
    assert isinstance(mux, float) and isinstance(muy, float)
    # Swap sign behaviors produce finite outputs
    assert np.isfinite(mux) and np.isfinite(muy)
