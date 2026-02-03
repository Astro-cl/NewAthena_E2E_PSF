import pytest

from main import compute_total_rot_polar


def test_polar_projection_unit_radial():
    # Single MM at x=1,y=0 -> radial unit vector (1,0)
    mm_to_pos = {1: 'A'}
    mm_config_map = {1: {'x_MM': 1.0, 'y_MM': 0.0, 'r_MM': 1.0}}
    # Supply gravity/cartesian rotations (gravity/thermal supply cartesian rotx/roty)
    alignment_by_pos = {}
    gravity_by_pos = {'A': {'d_grav_rotx': 1.0, 'd_grav_roty': 0.0}}
    thermal_by_pos = {}

    rotx, roty, rot_rad, rot_azi = compute_total_rot_polar(mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos)
    assert rotx['A'] == pytest.approx(1.0)
    assert roty['A'] == pytest.approx(0.0)
    assert rot_rad['A'] == pytest.approx(1.0)
    assert rot_azi['A'] == pytest.approx(0.0)


def test_polar_projection_unit_azimuthal():
    # Single MM at x=0,y=1 -> radial unit vector (0,1), azimuthal (-1,0)
    mm_to_pos = {2: 'B'}
    mm_config_map = {2: {'x_MM': 0.0, 'y_MM': 1.0, 'r_MM': 1.0}}
    alignment_by_pos = {}
    gravity_by_pos = {'B': {'d_grav_rotx': 0.0, 'd_grav_roty': 1.0}}
    thermal_by_pos = {}

    rotx, roty, rot_rad, rot_azi = compute_total_rot_polar(mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos)
    assert rotx['B'] == pytest.approx(0.0)
    assert roty['B'] == pytest.approx(1.0)
    # radial projection should be 1.0
    assert rot_rad['B'] == pytest.approx(1.0)
    # azimuthal projection should be 0.0 (since azimuthal unit is (-1,0) and vector is (0,1))
    assert rot_azi['B'] == pytest.approx(0.0)


def test_combined_sources():
    mm_to_pos = {1: 'A', 3: 'C'}
    mm_config_map = {
        1: {'x_MM': 1.0, 'y_MM': 0.0, 'r_MM': 1.0},
        3: {'x_MM': 0.0, 'y_MM': -2.0, 'r_MM': 2.0},
    }
    alignment_by_pos = {}
    gravity_by_pos = {'A': {'d_grav_rotx': 1.0, 'd_grav_roty': 0.0}}
    thermal_by_pos = {'C': {'d_therm_rotx': 1.0, 'd_therm_roty': 0.0}}

    rotx, roty, rot_rad, rot_azi = compute_total_rot_polar(mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos)
    # For pos 'A' total rotx = 1.0, roty = 0.0 -> radial 1.0, azimuth 0.0
    assert rotx['A'] == pytest.approx(1.0)
    assert roty['A'] == pytest.approx(0.0)
    assert rot_rad['A'] == pytest.approx(1.0)
    assert rot_azi['A'] == pytest.approx(0.0)
    # For pos 'C', only thermal applied on mm 3 at y negative
    assert rotx['C'] == pytest.approx(1.0)
    assert roty['C'] == pytest.approx(0.0)
    # radial unit vector is (0,-1) so radial projection = 0*1 + -1*0 = 0
    assert rot_rad['C'] == pytest.approx(0.0)
    # azimuthal = -rotx*u_rad_y + roty*u_rad_x = -1*(-1) + 0*0 = 1
    assert rot_azi['C'] == pytest.approx(1.0)
