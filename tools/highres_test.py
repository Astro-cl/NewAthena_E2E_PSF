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
print(base[['MM #','sigma_rad','sigma_azi']].head())

row = base.iloc[0]
mm = int(row['MM #'])
sigx = float(row['sigma_rad'])
sigy = float(row['sigma_azi'])
if abs(sigx - sigy) < 1e-12:
    sigma = sigx
else:
    sigma = np.sqrt(0.5*(sigx**2 + sigy**2))
hew_analytic = 2.0 * sigma * np.sqrt(2.0 * np.log(2.0))

n_r = 800
n_theta = 720
r_margin_factor = 6.0
max_sigma = max(sigx, sigy)
max_center = 0.0
r_max = max_center + r_margin_factor * max_sigma
if r_max <= 0:
    r_max = 1e-9

print(f"MM {mm}: sigx={sigx:.3e} m sigy={sigy:.3e} m | analytic HEW={hew_analytic*1e6:.6f} um")

start = time.time()
theta = np.linspace(0.0, 2.0 * np.pi, n_theta, endpoint=False)
r = np.linspace(0.0, r_max, n_r)
dtheta = theta[1] - theta[0]
dr = r[1] - r[0] if n_r > 1 else r_max
R, TH = np.meshgrid(r, theta)
Xp = R * np.cos(TH)
Yp = R * np.sin(TH)

Zp = gaussian_2d_rotated(
    Xp, Yp,
    mux=0.0, muy=0.0,
    sigmax=sigx, sigmay=sigy,
    theta=0.0,
    amplitude=1.0,
    normalize=True,
    degrees=True,
)

radial_energy = np.sum(Zp * R, axis=0) * dtheta
cumulative = np.cumsum(radial_energy * dr)
total = cumulative[-1] if cumulative.size else 1.0
frac = cumulative / total if total > 0 else cumulative
radius50 = np.interp(0.5, frac, r)
hew_sim = 2.0 * radius50
end = time.time()
print(f"High-res grid: n_r={n_r}, n_theta={n_theta}, r_margin_factor={r_margin_factor}, r_max={r_max:.6e} m")
print(f"Sim HEW={hew_sim*1e6:.6f} um | ratio sim/analytic={hew_sim/hew_analytic:.6f} | time={end-start:.2f}s")
