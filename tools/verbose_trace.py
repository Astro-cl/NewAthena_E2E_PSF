#!/usr/bin/env python3
import sys, os, argparse, pickle, json
from pathlib import Path
import numpy as np
import pandas as pd
sys.path.insert(0, os.getcwd())
import main

arcsec_to_m = 12 * np.pi / 180 / 3600

parser = argparse.ArgumentParser(description='Verbose trace: centroid and moments + plot_sum debug for a pickle')
parser.add_argument('pickle', help='Path to input pickle')
args = parser.parse_args()

p = Path(args.pickle)
if not p.exists():
    raise SystemExit(f"Pickle not found: {p}")

with p.open('rb') as fh:
    dfs = pickle.load(fh)

main.enable_inmemory_reads(dfs)
df = main.load_gaussians_from_excel('__INMEM__', sheet='MM_PSF')

# compute weights and centroid
w = df['weight'].to_numpy(dtype=float)
wsum = np.nansum(w) if w.size else 0.0
if (wsum and np.isfinite(wsum) and wsum > 0.0):
    wnorm = w/wsum
else:
    wnorm = np.ones_like(w) / max(1, len(w))

mux = df['mux'].to_numpy(dtype=float)
muy = df['muy'].to_numpy(dtype=float)
center_x = float((mux*wnorm).sum())
center_y = float((muy*wnorm).sum())

# weighted covariance
dx = mux - center_x
dy = muy - center_y
var_x = float((wnorm * dx * dx).sum())
var_y = float((wnorm * dy * dy).sum())
cov_xy = float((wnorm * dx * dy).sum())

# sigma stats
sig_rad = df['sigma_rad'].to_numpy(dtype=float)
sig_azi = df['sigma_azi'].to_numpy(dtype=float)
max_sig = float(max(np.nanmax(sig_rad), np.nanmax(sig_azi)))
mean_sig_rad = float(np.nanmean(sig_rad))
mean_sig_azi = float(np.nanmean(sig_azi))

# top MMs by weight and by sigma
idx_weight = np.argsort(-w)[:10]
idx_sigma = np.argsort(-np.maximum(sig_rad, sig_azi))[:10]

print(f"Pickle: {p}")
print(f"Total weight sum: {wsum:e}")
print(f"Centroid (mux,muy) [m]: ({center_x:.6e}, {center_y:.6e})")
print(f"Weighted variances [m^2]: var_x={var_x:.3e}, var_y={var_y:.3e}, cov_xy={cov_xy:.3e}")
print(f"Mean sigma_rad [arcsec]: {mean_sig_rad/arcsec_to_m:.6f}, mean sigma_azi [arcsec]: {mean_sig_azi/arcsec_to_m:.6f}")
print(f"Max sigma [arcsec]: {max_sig/arcsec_to_m:.6f}")
print("Top 10 by weight (MM #, weight):")
for i in idx_weight:
    mm = int(df.loc[i,'MM #']) if 'MM #' in df.columns else i
    print(f"  {mm}: {w[i]:.6e}")
print("Top 10 by sigma (MM #, sigma_arcsec):")
for i in idx_sigma:
    mm = int(df.loc[i,'MM #']) if 'MM #' in df.columns else i
    print(f"  {mm}: {max(sig_rad[i], sig_azi[i])/arcsec_to_m:.6f} arcsec")

# call plot_sum with debug
metrics = main.plot_sum(df, normalize=True, fast=True, return_metrics_only=True, debug=True)
print('\nMetrics JSON:')
print(json.dumps(metrics, indent=2))
