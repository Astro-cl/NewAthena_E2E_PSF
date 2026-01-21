import pandas as pd
import tempfile
import os
from main import load_aeff_weight_map


def test_load_aeff_weight_map_valid(tmp_path):
    # Create a minimal Excel file with A_eff sheet
    df = pd.DataFrame([
        ['MM #', 'A_eff'],
        [1, 0.123],
        [2, 0.456],
    ])
    p = tmp_path / 'test_aeff.xlsx'
    with pd.ExcelWriter(p, engine='openpyxl') as w:
        df.to_excel(w, sheet_name='A_eff', index=False, header=False)
    mapping = load_aeff_weight_map(str(p), sheet='A_eff')
    assert mapping[1] == 0.123
    assert mapping[2] == 0.456


def test_load_aeff_weight_map_invalid_missing_weight(tmp_path):
    df = pd.DataFrame([
        ['MM #', 'A_eff'],
        [1, 'not-a-number'],
    ])
    p = tmp_path / 'test_aeff2.xlsx'
    with pd.ExcelWriter(p, engine='openpyxl') as w:
        df.to_excel(w, sheet_name='A_eff', index=False, header=False)
    try:
        load_aeff_weight_map(str(p), sheet='A_eff')
        assert False, 'Expected ValueError for invalid weight'
    except ValueError:
        pass
