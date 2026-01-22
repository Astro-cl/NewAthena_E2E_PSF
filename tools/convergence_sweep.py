#!/usr/bin/env python3
import time, sys, os
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
from optimize_mm_rows import _load_base_params_from_workbook
from distributions_rotated import gaussian_2d_rotated
import numpy as np
import pandas as pd

file = 'Distributions/Test_Just_MM_8.xlsx'
print('Loading', file)
base = _load_base_params_from_workbook(file)
row = base.iloc[0]
mm = int(row['MM #'])
sigx = float(row['sigma_rad'])
sigy = float(row['sigma_azi'])
if abs(sigx - sigy) < 1e-12:
    sigma = sigx
else:
    sigma = np.sqrt(0.5*(sigx**2 + sigy**2))
hew_analytic = 2.0 * sigma * np.sqrt(2.0 * np.log(2.0))
print(f"MM {mm}: sigx={sigx:.3e} m sigy={sigy:.3e} m | analytic HEW={hew_analytic*1e6:.6f} um")

# Sweep n_r
n_theta_fixed = 360
r_margin_fixed = 5.0
n_r_list = [60, 120, 240, 480, 800]
print('\nSweep n_r (n_theta=360, r_margin=5):')
for n_r in n_r_list:
    start = time.time()
    n_theta = n_theta_fixed
    r_margin = r_margin_fixed
    max_sigma = max(sigx, sigy)
    r_max = r_margin * max_sigma
    if r_max <= 0:
        r_max = 1e-9
    theta = np.linspace(0.0, 2.0 * np.pi, n_theta, endpoint=False)
    r = np.linspace(0.0, r_max, n_r)
    dtheta = theta[1] - theta[0]
    dr = r[1] - r[0] if n_r > 1 else r_max
    R, TH = np.meshgrid(r, theta)
    Xp = R * np.cos(TH)
    Yp = R * np.sin(TH)
    Zp = gaussian_2d_rotated(Xp, Yp, mux=0.0, muy=0.0, sigmax=sigx, sigmay=sigy, theta=0.0, amplitude=1.0, normalize=True, degrees=True)
    radial_energy = np.sum(Zp * R, axis=0) * dtheta
    cumulative = np.cumsum(radial_energy * dr)
    total = cumulative[-1] if cumulative.size else 1.0
    frac = cumulative / total if total > 0 else cumulative
    radius50 = np.interp(0.5, frac, r)
    hew_sim = 2.0 * radius50
    end = time.time()
    print(f"n_r={n_r:4d} n_theta={n_theta:3d} r_margin={r_margin:.1f} | sim={hew_sim*1e6:9.6f} um | ratio={hew_sim/hew_analytic:.6f} | t={end-start:.3f}s")

# Sweep n_theta
n_r_fixed = 400
r_margin_fixed = 5.0
n_theta_list = [20, 40, 80, 160, 320, 720]
print('\nSweep n_theta (n_r=400, r_margin=5):')
for n_theta in n_theta_list:
    start = time.time()
    n_r = n_r_fixed
    r_margin = r_margin_fixed
    max_sigma = max(sigx, sigy)
    r_max = r_margin * max_sigma
    if r_max <= 0:
        r_max = 1e-9
    theta = np.linspace(0.0, 2.0 * np.pi, n_theta, endpoint=False)
    r = np.linspace(0.0, r_max, n_r)
    dtheta = theta[1] - theta[0]
    dr = r[1] - r[0] if n_r > 1 else r_max
    R, TH = np.meshgrid(r, theta)
    Xp = R * np.cos(TH)
    Yp = R * np.sin(TH)
    Zp = gaussian_2d_rotated(Xp, Yp, mux=0.0, muy=0.0, sigmax=sigx, sigmay=sigy, theta=0.0, amplitude=1.0, normalize=True, degrees=True)
    radial_energy = np.sum(Zp * R, axis=0) * dtheta
    cumulative = np.cumsum(radial_energy * dr)
    total = cumulative[-1] if cumulative.size else 1.0
    frac = cumulative / total if total > 0 else cumulative
    radius50 = np.interp(0.5, frac, r)
    hew_sim = 2.0 * radius50
    end = time.time()
    print(f"n_r={n_r:4d} n_theta={n_theta:3d} r_margin={r_margin:.1f} | sim={hew_sim*1e6:9.6f} um | ratio={hew_sim/hew_analytic:.6f} | t={end-start:.3f}s")

# Sweep r_margin
n_r_fixed = 400
n_theta_fixed = 360
r_margin_list = [2.0, 3.0, 4.0, 5.0, 6.0, 8.0]
print('\nSweep r_margin (n_r=400, n_theta=360):')
for r_margin in r_margin_list:
    start = time.time()
    n_r = n_r_fixed
    n_theta = n_theta_fixed
    max_sigma = max(sigx, sigy)
    r_max = r_margin * max_sigma
    if r_max <= 0:
        r_max = 1e-9
    theta = np.linspace(0.0, 2.0 * np.pi, n_theta, endpoint=False)
    r = np.linspace(0.0, r_max, n_r)
    dtheta = theta[1] - theta[0]
    dr = r[1] - r[0] if n_r > 1 else r_max
    R, TH = np.meshgrid(r, theta)
    Xp = R * np.cos(TH)
    Yp = R * np.sin(TH)
    Zp = gaussian_2d_rotated(Xp, Yp, mux=0.0, muy=0.0, sigmax=sigx, sigmay=sigy, theta=0.0, amplitude=1.0, normalize=True, degrees=True)
    radial_energy = np.sum(Zp * R, axis=0) * dtheta
    cumulative = np.cumsum(radial_energy * dr)
    total = cumulative[-1] if cumulative.size else 1.0
    frac = cumulative / total if total > 0 else cumulative
    radius50 = np.interp(0.5, frac, r)
    hew_sim = 2.0 * radius50
    end = time.time()
    print(f"n_r={n_r:4d} n_theta={n_theta:3d} r_margin={r_margin:.1f} | sim={hew_sim*1e6:9.6f} um | ratio={hew_sim/hew_analytic:.6f} | t={end-start:.3f}s")
