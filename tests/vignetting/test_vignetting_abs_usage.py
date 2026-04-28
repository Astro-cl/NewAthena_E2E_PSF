import numpy as np
from math import isclose


def test_vignetting_uses_absolute_rotation_for_azi():
    # simple vignette curve: factor decreases with increasing |rot| (arcsec)
    xs = np.array([0.0, 10.0, 20.0, 40.0])
    ys = np.array([1.0, 0.95, 0.9, 0.8])

    rot_neg = -20.0
    rot_pos = 20.0

    # expected behavior: use abs(rot) when interpolating
    val_neg_abs = float(np.interp(abs(rot_neg), xs, ys))
    val_pos_abs = float(np.interp(abs(rot_pos), xs, ys))

    assert isclose(val_neg_abs, val_pos_abs, rel_tol=1e-9)


def test_vignetting_uses_absolute_rotation_for_rad():
    xs = np.array([-10.0, 0.0, 10.0])
    ys = np.array([0.5, 1.0, 1.5])

    rot_neg = -7.5
    rot_pos = 7.5

    val_neg_abs = float(np.interp(abs(rot_neg), xs, ys))
    val_pos_abs = float(np.interp(abs(rot_pos), xs, ys))

    assert isclose(val_neg_abs, val_pos_abs, rel_tol=1e-9)
    assert isclose(val_neg_abs, 1.375, rel_tol=1e-9)
