import importlib
import tempfile
from pathlib import Path

import pandas as pd

from main import parse_multisheet_csv


def test_parse_multisheet_csv_multisheet(tmp_path: Path):
    content = (
        "# sheet: A\n"
        "col1,col2\n"
        "1,2\n\n"
        "# sheet: MM_PSF\n"
        "MM #,sigma_rad [arcsec]\n"
        "1,0.5\n"
    )
    p = tmp_path / "multi.csv"
    p.write_text(content, encoding="utf-8")

    sheets = parse_multisheet_csv(str(p))
    assert "A" in sheets and "MM_PSF" in sheets
    assert sheets["A"].shape[0] == 1
    assert sheets["MM_PSF"].shape[0] == 1
    assert float(sheets["MM_PSF"].iloc[0]["sigma_rad [arcsec]"]) == 0.5


def test_parse_standard_dist_spec_examples():
    mod = importlib.import_module("sensitivity.sensitivity_run")
    parse = getattr(mod, "_parse_standard_dist_spec")

    g = parse("gaussian(1.0,0.2)")
    assert isinstance(g, tuple) and g[0] == 'gaussian'
    assert abs(float(g[1]) - 1.0) < 1e-9
    assert abs(float(g[2]) - 0.2) < 1e-9

    u = parse("uniform(1.5,9)")
    assert isinstance(u, tuple) and u[0] == 'uniform'
    assert abs(float(u[1]) - 1.5) < 1e-9
    assert abs(float(u[2]) - 9.0) < 1e-9

    f = parse("5")
    assert isinstance(f, tuple) and f[0] == 'fixed'
    assert abs(float(f[1]) - 5.0) < 1e-9
