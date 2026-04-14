import importlib.util
from pathlib import Path
import textwrap


def load_main_module():
    repo_root = Path(__file__).resolve().parents[1]
    mod_path = repo_root / 'main.py'
    spec = importlib.util.spec_from_file_location('mainmod', str(mod_path))
    mod = importlib.util.module_from_spec(spec)
    import sys
    cwd_added = False
    try:
        if str(repo_root) not in sys.path:
            sys.path.insert(0, str(repo_root))
            cwd_added = True
        spec.loader.exec_module(mod)
    finally:
        if cwd_added:
            try:
                sys.path.remove(str(repo_root))
            except Exception:
                pass
    return mod


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

    main = load_main_module()
    sheets = main.parse_multisheet_csv(str(p))
    assert 'MM_PSF' in sheets
    assert 'A_eff' in sheets
    df_mm = sheets['MM_PSF']
    assert list(df_mm.columns) == ['a', 'b']
    assert int(df_mm.iloc[0, 0]) == 1


def load_sensitivity_module():
    repo_root = Path(__file__).resolve().parents[1]
    mod_path = repo_root / 'sensitivity' / 'sensitivity_run.py'
    spec = importlib.util.spec_from_file_location('sres', str(mod_path))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_parse_standard_dist_spec_basic():
    mod = load_sensitivity_module()
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
