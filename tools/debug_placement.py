#!/usr/bin/env python3
"""Debug placement for a single combo: print per-MM mapping, total weight,
and pre/post placement positions and sigma values.
"""
from pathlib import Path
import sys
import numpy as np
import pandas as pd
import math
import json
import time
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from tools.run_sensitivity import build_aeff_mapping, find_aeff_weights_for_choice, load_mm_row_map, apply_mm_psf_choice_to_df, load_standard_mm_psf_presets
from main import load_gaussians_from_excel, plot_sum
from optimize_mm_rows import load_all_sheets, _load_base_params_from_workbook, _load_position_deltas, _elliptical_place_mm_config, rebuild_df, compute_individual_mm_hew

BASE = ROOT / 'Distributions' / 'Test_Distribution.xlsx'

# Select a combo to debug (matches run_foreground first row)
AEFF_CHOICE = '1 keV [row1]'
MM_PSF_CHOICE = 'Fixed Sym Gaussian (sigma_rad=0.05 micron, sigma_azi=0.05 micron)'

print('Loading A_eff map and presets...')
aeff_map = build_aeff_mapping()
if 'mm_row_map' not in aeff_map:
    aeff_map['mm_row_map'] = load_mm_row_map(BASE)
standard = load_standard_mm_psf_presets(BASE)

# deterministic RNG for combo
import hashlib
h = hashlib.sha1(repr({'A_eff':AEFF_CHOICE,'MM_PSF':MM_PSF_CHOICE}).encode('utf8')).hexdigest()
seed = int(h[:16],16) % (2**32)
rng = np.random.default_rng(seed)

print('Loading runtime DF...')
df = load_gaussians_from_excel(str(BASE), sheet='MM_PSF')
# apply A_eff mapping
mapping = find_aeff_weights_for_choice(AEFF_CHOICE, aeff_map)

df['weight'] = df['MM #'].astype(int).map(mapping)
# apply MM_PSF
apply_mm_psf_choice_to_df(df, MM_PSF_CHOICE, aeff_map, rng, standard_presets=standard)

print('Total mapped weight:', float(df['weight'].sum()))
print('Top 10 MMs by weight:')
print(df[['MM #','weight']].sort_values('weight',ascending=False).head(10).to_string(index=False))

# compute pre-placement metrics
pre = plot_sum(df, normalize=True, fast=True, df_optimized=None, return_metrics_only=True, debug=True)
print('Pre-placement metrics:', pre)

# Load base params and run in-process placement
sheets = load_all_sheets(str(BASE))
mm_config = sheets.get('MM configuration')
base_params = _load_base_params_from_workbook(str(BASE))
# override base_params weights with mapping
mapped = base_params['MM #'].astype(int).map(mapping).fillna(0.0)
base_params['weight'] = mapped
# Do not copy runtime per-MM fields here; rebuild_df preserves runtime values.
print('Base params total weight (after mapping):', float(base_params['weight'].sum()))

# Apply the MM_PSF choice to base_params as well so placement uses the same
# per-MM sigmas/alphas as the runtime input DF. This mirrors the behavior
# in the sensitivity runner where the choice is applied before placement.
try:
    apply_mm_psf_choice_to_df(base_params, MM_PSF_CHOICE, aeff_map, rng, standard_presets=standard)
except Exception as e:
    print('Warning: could not apply MM_PSF choice to base_params:', e)

alignment_by_pos, gravity_by_pos, thermal_by_pos = _load_position_deltas(str(BASE))
placed_mm = _elliptical_place_mm_config(mm_config, base_params, alignment_by_pos=alignment_by_pos, gravity_by_pos=gravity_by_pos, thermal_by_pos=thermal_by_pos, seed=int(seed))
final_df = rebuild_df(base_params, placed_mm)

# compute post-placement metrics
post = plot_sum(df, normalize=True, fast=True, df_optimized=final_df, return_metrics_only=True, debug=True)
print('Post-placement metrics:', post)

print('\nRecomputing metrics with fast=False for higher accuracy...')
post_high = plot_sum(df, normalize=True, fast=False, df_optimized=final_df, return_metrics_only=True, debug=True)
print('Post-placement metrics (fast=False):', post_high)

# Extra diagnostics: build summed optimized grid and report centroid and peak
from distributions_rotated import gaussian_2d_rotated
print('\nBuilding sampled grid for optimized DF to inspect Z_opt...')
opt_df = final_df.copy()
opt_mux = opt_df['mux'].to_numpy(dtype=float)
opt_muy = opt_df['muy'].to_numpy(dtype=float)
opt_sigx = opt_df['sigmax'].to_numpy(dtype=float)
opt_sigy = opt_df['sigmay'].to_numpy(dtype=float)
opt_w = opt_df['weight'].to_numpy(dtype=float)

# grid extents
cx = float((opt_df['mux'] * opt_df['weight']).sum() / opt_df['weight'].sum()) if opt_df['weight'].sum() > 0 else 0.0
cy = float((opt_df['muy'] * opt_df['weight']).sum() / opt_df['weight'].sum()) if opt_df['weight'].sum() > 0 else 0.0
max_sigma = max(opt_df['sigmax'].max(), opt_df['sigmay'].max())
max_center_dist = np.sqrt((opt_df['mux'] - cx) ** 2 + (opt_df['muy'] - cy) ** 2).max()
rmax = max_center_dist + 5.0 * max_sigma
if rmax <= 0:
    rmax = 1e-6
nx = ny = 201
xs = np.linspace(cx - rmax, cx + rmax, nx)
ys = np.linspace(cy - rmax, cy + rmax, ny)
X, Y = np.meshgrid(xs, ys)
Z = np.zeros_like(X, dtype=float)
for i in range(len(opt_mux)):
    Z += opt_w[i] * gaussian_2d_rotated(X, Y, mux=opt_mux[i], muy=opt_muy[i], sigmax=opt_sigx[i], sigmay=opt_sigy[i], theta=0.0, normalize=True, degrees=True)

# compute peak location and centroid
imax = np.unravel_index(np.nanargmax(Z), Z.shape)
peak_x = X[imax]
peak_y = Y[imax]
centroid_x = (X * Z).sum() / Z.sum() if Z.sum() > 0 else np.nan
centroid_y = (Y * Z).sum() / Z.sum() if Z.sum() > 0 else np.nan
print(f' Z_opt peak at (m): {peak_x:.6e}, {peak_y:.6e}    centroid (m): {centroid_x:.6e}, {centroid_y:.6e}')
print(f' opt center-weighted average (m): {cx:.6e}, {cy:.6e}  total weight: {opt_df["weight"].sum():.6e}')

# compare per-MM positions
print('\nPer-MM comparison (top 20 by weight):')
merged = df.set_index('MM #')[['weight','sigmax','sigmay']].join(final_df.set_index('MM #')[['mux','muy']], how='left')
# original center estimates (from base_params) available in base_params 'mux','muy' columns if present
if 'mux' in base_params.columns:
    merged['base_mux'] = base_params.set_index('MM #')['mux']
    merged['base_muy'] = base_params.set_index('MM #')['muy']

# Also include optimized mux/muy from final_df
merged = merged.reset_index()
merged_sorted = merged.sort_values('weight', ascending=False).head(20)
print(merged_sorted.to_string(index=False))

# Save diagnostic dumps for further inspection
OUT_DIR = ROOT / 'Figures'
OUT_DIR.mkdir(parents=True, exist_ok=True)
ts = int(time.time())
csv_base = OUT_DIR / f'debug_combo_{ts}'
try:
    final_df.to_csv(str(csv_base) + '_final_df.csv', index=False)
    base_params.to_csv(str(csv_base) + '_base_params.csv', index=False)
    df.to_csv(str(csv_base) + '_input_df.csv', index=False)
    merged_sorted.to_csv(str(csv_base) + '_merged_top20.csv', index=False)
    # save sampled grid
    np.savez(str(csv_base) + '_Z_opt.npz', X=X, Y=Y, Z=Z, peak_x=peak_x, peak_y=peak_y, centroid_x=centroid_x, centroid_y=centroid_y)
    # save metrics
    metrics = {'pre': pre, 'post_fast': post, 'post_slow': post_high}
    with open(str(csv_base) + '_metrics.json', 'w') as fh:
        json.dump({k: (v if not isinstance(v, np.generic) else float(v)) for k,v in metrics.items()}, fh, default=str)
    print('\nSaved diagnostic dumps to', OUT_DIR)
except Exception as e:
    print('Failed to save diagnostic dumps:', e)

# compute per-MM HEW contributions using compute_individual_mm_hew if available
try:
    hew_contribs = []
    for _, row in merged_sorted.iterrows():
        mm = int(row['MM #'])
        # build param row for compute_individual_mm_hew: expect base_params DataFrame row and mm param series
        base_row = base_params[base_params['MM #']==mm]
        if base_row.empty:
            continue
        # compute impact (some helper functions expect full df) — call compute_individual_mm_hew
        hew = compute_individual_mm_hew(final_df, mm)
        hew_contribs.append((mm, hew))
    print('\nPer-MM HEW estimates (approx):')
    for mm, hval in hew_contribs:
        print(f' MM {mm}: HEW ~ {hval}')
except Exception as e:
    print('Could not compute per-MM HEW contributions:', e)

print('\nDone')
