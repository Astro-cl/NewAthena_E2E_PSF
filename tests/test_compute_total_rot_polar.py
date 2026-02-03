import os
import sys

# Ensure repository root is on sys.path so tests can import `main` directly
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pytest

from main import compute_total_rot_polar


def test_compute_total_rot_polar_projection_sum():
    # Setup: one MM mapped to position 1 with radial unit vector along x (ux=1, uy=0)
    mm_to_pos = {100: 1}
    mm_config_map = {100: {'x_MM': 1.0, 'y_MM': 0.0, 'r_MM': 1.0}}

    # Alignment provides direct polar contributions (rotazi, rotrad)
    alignment_by_pos = {1: {'d_align_rotazi': 2.0, 'd_align_rotrad': 5.0}}

    # Gravity provides both rotx/roty (which should be projected) and direct polar
    gravity_by_pos = {
        1: {
            'd_grav_rotx': 10.0,  # arcsec
            'd_grav_roty': 0.0,
            'd_grav_rotazi': 3.0,
            'd_grav_rotrad': 4.0,
        }
    }

    thermal_by_pos = {}

    rotx_map, roty_map, rot_rad_map, rot_azi_map = compute_total_rot_polar(
        mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos
    )

    # Projected contribution: rotx * ux + roty * uy = 10*1 + 0*0 = 10
    # Direct polar total: alignment_rotrad (5) + gravity_rotrad (4) = 9
    # Expected total rotrad = 10 + 9 = 19
    assert pytest.approx(rotx_map[1], rel=1e-12) == 10.0
    assert pytest.approx(roty_map[1], rel=1e-12) == 0.0
    assert pytest.approx(rot_rad_map[1], rel=1e-12) == 19.0

    # For rot_azi: projected (-rotx*uy + roty*ux) = 0; direct rotazi = 2 + 3 = 5
    assert pytest.approx(rot_azi_map[1], rel=1e-12) == 5.0
