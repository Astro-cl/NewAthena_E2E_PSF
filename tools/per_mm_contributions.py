#!/usr/bin/env python3
import sys, os, argparse, pickle, json
from pathlib import Path
import numpy as np
import pandas as pd

sys.path.insert(0, os.getcwd())
import main
from distributions_rotated import gaussian_2d_rotated, pseudo_voigt_2d_rotated, load_psf_matrix_excel, eval_psf_matrix_rotated

arcsec_to_m = 12 * np.pi / 180 / 3600

parser = argparse.ArgumentParser(description='Compute per-MM contribution to HEW/EEF for a pickle')
parser.add_argument('pickle', help='input pickle path')
parser.add_argument('--n-r', type=int, default=200)
parser.add_argument('--n-theta', type=int, default=180)
parser.add_argument('--r-margin-factor', type=float, default=5.0)
parser.add_argument('--top', type=int, default=10, help='Show top N contributors')
args = parser.parse_args()

p = Path(args.pickle)
if not p.exists():
    raise SystemExit(f"Pickle not found: {p}")

with p.open('rb') as fh:
    dfs = pickle.load(fh)

main.enable_inmemory_reads(dfs)
df = main.load_gaussians_from_excel('__INMEM__', sheet='MM_PSF')

# Prepare arrays mimicking main.plot_sum
mux = df['mux'].to_numpy(dtype=float)
muy = df['muy'].to_numpy(dtype=float)
sigx = df['sigma_rad'].to_numpy(dtype=float)
sigy = df['sigma_azi'].to_numpy(dtype=float)
theta = df['theta_degrees'].to_numpy(dtype=float)

# distribution names handling
dist_raw = df.get('distribution', pd.Series(['gaussian'] * len(df))).astype(str).to_numpy(copy=False)
dist = pd.Series(dist_raw).astype(str).str.lower().to_numpy(copy=False)
alpha_azi = pd.to_numeric(df.get('alpha_azi', pd.Series([0.5] * len(df))), errors='coerce').fillna(0.5).to_numpy(dtype=float)
alpha_rad = pd.to_numeric(df.get('alpha_rad', pd.Series([0.5] * len(df))), errors='coerce').fillna(0.5).to_numpy(dtype=float)

# weights normalized
if 'weight' in df.columns:
    wtmp = df['weight'].to_numpy(dtype=float)
    wsum = float(np.nansum(wtmp)) if wtmp.size else 0.0
    weight_arr = (wtmp / wsum) if (wsum and np.isfinite(wsum) and wsum > 0.0) else np.ones(len(df), dtype=float)
else:
    weight_arr = np.ones(len(df), dtype=float)

# Prevent zero sigmas
MIN_SIG_M = 1e-9
sigx = np.maximum(sigx, MIN_SIG_M)
sigy = np.maximum(sigy, MIN_SIG_M)

# Custom PSF cache
workbook_path = df.attrs.get('workbook_path', None)
custom_cache = {}
if workbook_path is not None:
    unique_names = sorted({str(n).strip() for n, dl in zip(dist_raw, dist) if str(n).strip() and dl not in {'gaussian','pseudo-voigt','voigt'}})
    for name in unique_names:
        pth = main._resolve_custom_psf_path(workbook_path, name)
        if not pth:
            continue
        try:
            x_psf, y_psf, f_psf = load_psf_matrix_excel(pth, arcsec_to_m=arcsec_to_m)
            custom_cache[name] = (x_psf, y_psf, f_psf)
        except Exception:
            pass

# compute r_max same logic
max_sigma = max(float(sigx.max()), float(sigy.max()))
if 'custom_sigma_hint' in locals() and custom_cache:
    max_sigma = max(max_sigma, max([max(np.max(np.abs(x)), np.max(np.abs(y)))/3.0 for (x,y,f) in custom_cache.values()]))
max_center_dist = np.sqrt((mux - 0.0)**2 + (muy - 0.0)**2).max() if len(mux)>0 else 0.0
r_max = max_center_dist + args.r_margin_factor * max_sigma
if r_max <= 0:
    r_max = 1e-6

# build polar grid
theta_vals = np.linspace(0.0, 2.0*np.pi, args.n_theta, endpoint=False)
r_vals = np.linspace(0.0, r_max, args.n_r)
dtheta = theta_vals[1]-theta_vals[0] if args.n_theta>1 else 2*np.pi
dr = r_vals[1]-r_vals[0] if args.n_r>1 else r_max
R, TH = np.meshgrid(r_vals, theta_vals)
X = R*np.cos(TH)
Y = R*np.sin(TH)

# Evaluate per-MM PSF at grid points
n = len(mux)
Z_total = np.zeros_like(X, dtype=float)
Z_per_mm = np.zeros((n,) + X.shape, dtype=float)
for i in range(n):
    di = dist[i]
    if di in ['pseudo-voigt','voigt']:
        # positional args: azi, rad, muazi, murad, sigmaazi, sigmarad, theta
        Zi = pseudo_voigt_2d_rotated(X, Y, mux[i], muy[i], sigx[i], sigy[i], theta[i], alphaazi=alpha_azi[i], alpharad=alpha_rad[i])
    elif di == 'gaussian':
        # gaussian_2d_rotated(x, y, mux, muy, sigmax, sigmay, theta)
        Zi = gaussian_2d_rotated(X, Y, mux[i], muy[i], sigx[i], sigy[i], theta[i])
    else:
        # custom PSF
        name = str(dist_raw[i]).strip()
        if name in custom_cache:
            x_psf, y_psf, f_psf = custom_cache[name]
            Zi = eval_psf_matrix_rotated(X, Y, mux=mux[i], muy=muy[i], theta_deg=theta[i], x_axis=x_psf, y_axis=y_psf, flux=f_psf)
        else:
            # fallback zero
            Zi = np.zeros_like(X, dtype=float)
    Zi = Zi * weight_arr[i]
    Z_per_mm[i] = Zi
    Z_total += Zi

# radial integration
radial_energy_total = np.sum(Z_total * R, axis=0) * dtheta
cumulative_total = np.cumsum(radial_energy_total * dr)
total_energy = cumulative_total[-1] if cumulative_total.size else 1.0

# helper to interp cumulative at radius
from numpy import interp

def cum_at_radius(cumulative, r_vals, r):
    if r <= r_vals[0]:
        return cumulative[0]
    if r >= r_vals[-1]:
        return cumulative[-1]
    return float(interp(r, r_vals, cumulative))

# find radii for 50% and 90%
from bisect import bisect_left

def radius_for_fraction(cum, r_vals, target):
    frac = cum/ (cum[-1] if cum[-1]>0 else 1.0)
    idx = bisect_left(frac, target)
    if idx==0:
        return r_vals[0]
    if idx>=len(frac):
        return r_vals[-1]
    # linear interp
    f0, f1 = frac[idx-1], frac[idx]
    r0, r1 = r_vals[idx-1], r_vals[idx]
    if f1==f0:
        return r0
    return r0 + (target - f0)*(r1-r0)/(f1-f0)

r_50 = radius_for_fraction(cumulative_total, r_vals, 0.5)
r_90 = radius_for_fraction(cumulative_total, r_vals, 0.9)

# compute per-mm contribution at those radii
contribs = []
for i in range(n):
    rad_energy_i = np.sum(Z_per_mm[i] * R, axis=0) * dtheta
    cum_i = np.cumsum(rad_energy_i * dr)
    c50 = cum_at_radius(cum_i, r_vals, r_50)
    c90 = cum_at_radius(cum_i, r_vals, r_90)
    contribs.append({'index': i, 'MM #': int(df.loc[i,'MM #']) if 'MM #' in df.columns else i, 'weight': float(weight_arr[i]), 'c50': float(c50/total_energy), 'c90': float(c90/total_energy)})

contribs_sorted_50 = sorted(contribs, key=lambda x: x['c50'], reverse=True)
contribs_sorted_90 = sorted(contribs, key=lambda x: x['c90'], reverse=True)

out = {
    'pickle': str(p),
    'n_total': n,
    'total_energy': float(total_energy),
    'r_50_m': float(r_50),
    'r_90_m': float(r_90),
    'r_50_arcsec': float(r_50/arcsec_to_m),
    'r_90_arcsec': float(r_90/arcsec_to_m),
    'top_contributors_50': contribs_sorted_50[:args.top],
    'top_contributors_90': contribs_sorted_90[:args.top],
}

print(json.dumps(out, indent=2))
