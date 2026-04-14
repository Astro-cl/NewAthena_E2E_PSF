import importlib
from pathlib import Path
import textwrap


def test_parse_multisheet_csv_two_sheets(tmp_path):
    content = textwrap.dedent("""
    # sheet: MM_PSF
    a,b
    1,2

    # sheet: A_eff
    MM #,weight
    1,0.5
    """)
    p = tmp_path / "multi.csv"
    p.write_text(content, encoding='utf-8')

    # Import the repository `main` module directly
    main = importlib.import_module('main')
    sheets = main.parse_multisheet_csv(str(p))
    assert 'MM_PSF' in sheets
    assert 'A_eff' in sheets
    df_mm = sheets['MM_PSF']
    assert list(df_mm.columns) == ['a', 'b']
    assert int(df_mm.iloc[0, 0]) == 1


def test_parse_standard_dist_spec_basic():
    mod = importlib.import_module('sensitivity.sensitivity_run')
    fn = getattr(mod, '_parse_standard_dist_spec')

    res = fn('gaussian(0,12/3)')
    # gaussian returns ('gaussian', mean, sigma)
    assert isinstance(res, tuple)
    assert res[0] == 'gaussian'
    assert abs(res[1] - 0.0) < 1e-9
    assert abs(res[2] - 4.0) < 1e-9

    res2 = fn('1.23')
    assert res2[0] == 'fixed'
    assert abs(res2[1] - 1.23) < 1e-9
