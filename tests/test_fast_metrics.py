import math
import pytest
import sys, os
# ensure repo root is on path for test discovery when running pytest from any CWD
repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if repo_root not in sys.path:
    sys.path.insert(0, repo_root)
from main import load_gaussians_from_excel, plot_sum


def test_fast_metrics_matches_slow_tolerance():
    """Fast approximate metrics (fast-metrics) should match slow grid metrics
    within a reasonable tolerance for typical workbooks.

    This test compares HEW and EEF90 diameters (arcsec) between the
    analytical fast path and the reference slow polar-integration path.
    """
    path = 'Distributions/NewTest_Distribution.xlsx'

    # slow (reference) path
    df_ref = load_gaussians_from_excel(path, fast_metrics=False)
    ref = plot_sum(df_ref, return_metrics_only=True, fast=False)

    # fast approximate path
    df_fast = load_gaussians_from_excel(path, fast_metrics=True)
    fast = plot_sum(df_fast, return_metrics_only=True, fast=True)

    # keys to compare
    keys = ['hew_best_arcsec', 'hew_origin_arcsec', 'eef90_best_arcsec', 'eef90_origin_arcsec']

    for k in keys:
        rv = ref.get(k)
        fv = fast.get(k)
        assert rv is not None and fv is not None, f"Missing metric {k}: ref={rv} fast={fv}"
        # allow 15% relative error or 0.5 arcsec absolute, whichever is larger
        tol_rel = 0.15
        tol_abs = 0.5
        allowed = max(tol_abs, abs(rv) * tol_rel)
        diff = abs(rv - fv)
        assert diff <= allowed, f"Metric {k} differs too much: ref={rv} fast={fv} diff={diff} allowed={allowed}"
