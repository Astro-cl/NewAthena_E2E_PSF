import numpy as np
from distributions_rotated import _sort_axis_and_reorder, gaussian_2d_rotated, pseudo_voigt_2d_rotated


def test_sort_axis_and_reorder():
    axis = np.array([3.0, 1.0, 2.0])
    data = np.arange(9).reshape(3, 3)
    a_sorted, d_sorted = _sort_axis_and_reorder(axis, data, axis_is_x=True)
    assert np.allclose(a_sorted, np.array([1.0, 2.0, 3.0]))
    # Column order should follow axis order
    assert d_sorted.shape == data.shape
    assert list(d_sorted[:, 0]) == list(data[:, 1])


def test_gaussian_center_value_normalized():
    x = np.array([[0.0]])
    y = np.array([[0.0]])
    val = gaussian_2d_rotated(x, y, mux=0.0, muy=0.0, sigmax=1.0, sigmay=1.0, theta=0.0, amplitude=1.0, normalize=True)
    # For normalized Gaussian with sigma=1, value at center should equal 1/(2πσxσy)
    expected = 1.0 / (2.0 * np.pi * 1.0 * 1.0)
    assert np.allclose(val, expected)


def test_pseudo_voigt_basic_properties():
    # 2x2 grid
    x = np.array([[0.0, 1.0], [0.0, 1.0]])
    y = np.array([[0.0, 0.0], [1.0, 1.0]])
    pv = pseudo_voigt_2d_rotated(x, y, muazi=0.0, murad=0.0, sigmaazi=1.0, sigmarad=1.0, theta=0.0, eta=0.3, amplitude=2.0, normalize=False)
    assert pv.shape == x.shape
    assert np.all(np.isfinite(pv))
    assert np.all(pv >= 0.0)
