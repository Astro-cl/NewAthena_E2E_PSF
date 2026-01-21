import pandas as pd
import numpy as np
from gui_distributions import generate_values, generate_data_from_distributions, ExtendedGUI


def test_generate_values_fixed_and_uniform():
    a = 3.5
    arr = generate_values('fixed', a, 0, 5)
    assert len(arr) == 5
    assert all(np.isclose(arr, float(a)))

    arr2 = generate_values('uniform', 1.0, 2.0, 1000)
    assert arr2.min() >= 1.0 and arr2.max() <= 2.0


def test_generate_data_from_distributions_and_sigma_positive():
    params = {
        'sigma_rad [arcsec]': ('gaussian', 1.0, 0.1),
        'm_rad [arcsec]': ('fixed', 0.0, 0.0),
        'sigma_azi [arcsec]': ('fixed', 2.0, 0.0),
        'alpha_rad': ('fixed', 0.5, 0.0)
    }
    config = {
        'params': ['m_rad [arcsec]', 'm_azi [arcsec]', 'sigma_rad [arcsec]', 'sigma_azi [arcsec]'],
        'alpha_params': ['alpha_rad', 'alpha_azi']
    }
    df = generate_data_from_distributions(params, num_mm=10, data_type_config=config)
    assert 'sigma_rad [arcsec]' in df.columns
    assert (df['sigma_rad [arcsec]'] > 0).all()


def test_safe_eval_and_parse_spec():
    eg = object.__new__(ExtendedGUI)
    # _safe_eval_numeric_expr
    val = ExtendedGUI._safe_eval_numeric_expr(eg, '110%')
    assert np.isclose(val, 1.10)
    val2 = ExtendedGUI._safe_eval_numeric_expr(eg, '12/3')
    assert np.isclose(val2, 4.0)

    # _parse_standard_dist_spec
    kind, a, b = ExtendedGUI._parse_standard_dist_spec(eg, 'gaussian(3.0, 0.5)')
    assert kind == 'gaussian' and np.isclose(a, 3.0)


def test_evaluate_aeff_preset_for_mm_simple():
    eg = object.__new__(ExtendedGUI)
    # Build minimal aeff_raw_df where row index equals MM#, columns A,B,C... representing letters
    data = [
        ['MM #', 'A_eff', '', 'D', 'E'],
        [1, 0.1, None, None, None],
        [2, 0.2, None, None, None],
    ]
    eg.aeff_raw_df = pd.DataFrame(data)
    # Test simple 'J' style where J maps to column letter 'J' -> but our small df doesn't have J.
    # Instead, directly test _value_from_column_letter and _get_aeff_row_for_mm by setting up an accessible layout.
    # Create a layout where columns A..E exist and row 1 contains values.
    df2 = pd.DataFrame([
        ['MM #','A','B','C','D','E'],
        [1, 10, 11, 12, 13, 14],
        [2, 20, 21, 22, 23, 24]
    ])
    eg.aeff_raw_df = df2
    # _get_aeff_row_for_mm should find row index 1 for mm=1
    r = ExtendedGUI._get_aeff_row_for_mm(eg, 1)
    assert r == 1
    v = ExtendedGUI._value_from_column_letter(eg, 1, 'B')
    assert v == 11.0
    # Test expression evaluation: 'B+gaussian(0,0)' simplified via _evaluate_aeff_preset_for_mm
    # We'll add a simple preset expression and call _evaluate_aeff_preset_for_mm
    eg._get_aeff_row_for_mm = lambda mm: 1
    eg._value_from_column_letter = lambda row, col: 11.0 if col == 'B' else 0.0
    val = ExtendedGUI._evaluate_aeff_preset_for_mm(eg, 1, 'B')
    assert np.isclose(val, 11.0)
