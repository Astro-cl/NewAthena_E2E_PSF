import json
import subprocess
import math
from main import compute_hew_eef_metrics


def test_compute_hew_eef_metrics_keys():
    metrics = compute_hew_eef_metrics(file='Distributions/Test_Distribution.xlsx', sheet='MM_PSF', normalize=True, fast=True)
    assert isinstance(metrics, dict)
    required_keys = [
        'hew_origin_arcsec', 'hew_best_arcsec', 'eef90_origin_arcsec', 'eef90_best_arcsec',
        'hew_x_arcsec', 'hew_y_arcsec', 'hew_opt_x_arcsec', 'hew_opt_y_arcsec',
        'hew_opt_arcsec', 'eef90_opt_arcsec'
    ]
    for k in required_keys:
        assert k in metrics

    for k in ['hew_origin_arcsec', 'hew_best_arcsec', 'eef90_origin_arcsec', 'eef90_best_arcsec', 'hew_x_arcsec', 'hew_y_arcsec']:
        v = metrics[k]
        assert v is not None
        assert isinstance(v, (int, float))
        assert math.isfinite(v)


def test_cli_return_metrics_only_json():
    res = subprocess.run(
        ['python3', 'main.py', '-f', 'Distributions/Test_Distribution.xlsx', '--return_metrics_only'],
        capture_output=True,
        text=True,
        check=True,
    )
    out = res.stdout.strip()
    idx = out.find('{')
    assert idx != -1, 'No JSON object found in CLI output'
    json_str = out[idx:]
    metrics = json.loads(json_str)
    assert isinstance(metrics, dict)
    assert 'hew_best_arcsec' in metrics
