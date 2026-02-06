import pytest

from main import compute_total_rot_polar


def test_projection_case1_nontrivial_geometry():
    # MM at (x,y) with given radius -> ux,uy as in real workbook
    mm_to_pos = {1: 1}
    mm_config_map = {
        1: {
            'x_MM': 0.0434013330,
            'y_MM': 0.2711059987,
            'r_MM': 0.2745580781,
        }
    }
    alignment_by_pos = {}
    gravity_by_pos = {1: {'d_grav_rotx': 120.0, 'd_grav_roty': 0.0}}
    thermal_by_pos = {}

    rotx, roty, rot_rad, rot_azi = compute_total_rot_polar(
        mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos
    )

    ux = mm_config_map[1]['x_MM'] / mm_config_map[1]['r_MM']
    uy = mm_config_map[1]['y_MM'] / mm_config_map[1]['r_MM']

    expected_proj_rotrad = 120.0 * ux + 0.0 * uy
    expected_proj_rotazi = -120.0 * uy + 0.0 * ux

    assert rotx[1] == pytest.approx(120.0)
    assert roty[1] == pytest.approx(0.0)
    assert rot_rad[1] == pytest.approx(expected_proj_rotrad)
    assert rot_azi[1] == pytest.approx(expected_proj_rotazi)


def test_projection_case2_axis_aligned():
    # MM along +y axis: ux=0, uy=1 => rotx projects to rotazi with negative sign
    mm_to_pos = {2: 2}
    mm_config_map = {2: {'x_MM': 0.0, 'y_MM': 1.0, 'r_MM': 1.0}}
    alignment_by_pos = {}
    gravity_by_pos = {2: {'d_grav_rotx': 120.0, 'd_grav_roty': 0.0}}
    thermal_by_pos = {}

    rotx, roty, rot_rad, rot_azi = compute_total_rot_polar(
        mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos
    )

    assert rotx[2] == pytest.approx(120.0)
    assert roty[2] == pytest.approx(0.0)
    assert rot_rad[2] == pytest.approx(0.0)
    assert rot_azi[2] == pytest.approx(-120.0)
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
