import importlib


def test_compute_aeff_values_importable():
    """Module should be importable and expose `compute_aeff_values` function."""
    mod = importlib.import_module('compute_aeff_values')
    assert hasattr(mod, 'compute_aeff_values')
    assert callable(mod.compute_aeff_values)
