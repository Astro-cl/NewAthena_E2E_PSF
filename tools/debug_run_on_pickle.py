#!/usr/bin/env python3
import argparse, pickle, json
from pathlib import Path

import numpy as np
import pandas as pd

import sys, os
sys.path.insert(0, os.getcwd())
import main

arcsec_to_m = 12 * np.pi / 180 / 3600

parser = argparse.ArgumentParser(description='Debug run main.plot_sum on an input pickle with per-MM stats')
parser.add_argument('pickle', help='Path to input .pkl produced by sensitivity_run')
parser.add_argument('--sheet', default='MM_PSF')
parser.add_argument('--metrics-nr-final', type=int, default=50)
parser.add_argument('--metrics-ntheta-final', type=int, default=12)
parser.add_argument('--metrics-r-margin', type=float, default=6.0)
args = parser.parse_args()

p = Path(args.pickle)
if not p.exists():
    raise SystemExit(f"Pickle not found: {p}")

with p.open('rb') as fh:
    dfs = pickle.load(fh)

# Enable in-memory reads and load gaussians
main.enable_inmemory_reads(dfs)
df = main.load_gaussians_from_excel('__INMEM__', sheet=args.sheet)

# Print basic per-MM stats (convert sigmas back to arcsec)
sig_rad_m = df['sigma_rad'].to_numpy(dtype=float)
sig_azi_m = df['sigma_azi'].to_numpy(dtype=float)
alpha_rad = df.get('alpha_rad', None)
alpha_azi = df.get('alpha_azi', None)

def stats(arr):
    arr = np.asarray(arr, dtype=float)
    if arr.size == 0:
        return {'count':0}
    return {'count':int(arr.size), 'mean':float(np.nanmean(arr)), 'std':float(np.nanstd(arr)), 'min':float(np.nanmin(arr)), 'max':float(np.nanmax(arr))}

print('Per-MM sigma_rad [arcsec]:', json.dumps(stats(sig_rad_m/arcsec_to_m)))
print('Per-MM sigma_azi [arcsec]:', json.dumps(stats(sig_azi_m/arcsec_to_m)))
if alpha_rad is not None:
    print('alpha_rad count (non-null):', int((~pd.isna(alpha_rad)).sum()))
if alpha_azi is not None:
    print('alpha_azi count (non-null):', int((~pd.isna(alpha_azi)).sum()))

# Print first 100 sigma_rad values (arcsec) for quick inspection
vals = (sig_rad_m/arcsec_to_m).tolist()
print('First 100 sigma_rad [arcsec]:')
print(', '.join(f"{v:.6f}" for v in vals[:100]))

# Run plot_sum with return_metrics_only and debug=True
metrics = main.plot_sum(
    df,
    normalize=True,
    fast=True,
    return_metrics_only=True,
    debug=True,
    metrics_n_r_final=args.metrics_nr_final,
    metrics_n_theta_final=args.metrics_ntheta_final,
    metrics_r_margin=args.metrics_r_margin,
)

print('\nMetrics JSON:')
print(json.dumps(metrics, indent=2))
