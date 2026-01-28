from pathlib import Path
import shutil
import numpy as np
import pandas as pd
import subprocess
import json

# Target workbook that contains the 10% preset
target = Path('sensitivity/input/20260123T162222Z_2_MM_PSFsigma_azi0.05_arcsec_Alignment0.0_Gravity_offload0.0_Thermal0.0.xlsx')
if not target.exists():
    print('TARGET_NOT_FOUND', target)
    raise SystemExit(1)
new = target.with_name(target.stem + '_fixed.xlsx')
shutil.copy2(target, new)
print('copied->', new.name)
# load MM_PSF
try:
    df = pd.read_excel(new, sheet_name='MM_PSF', engine='openpyxl')
except Exception as e:
    print('read error', e)
    raise
# parse numeric template columns
sr = pd.to_numeric(df.get('sigma_rad_', None), errors='coerce')
sa = pd.to_numeric(df.get('sigma_azi_', None), errors='coerce')
if sr.isna().all():
    print('no numeric sigma_rad_ template found; abort')
    raise SystemExit(1)
# sample with pct=0.10
pct = 0.10
rng = np.random.default_rng(42)
mu_r = sr.fillna(0.0).to_numpy()
if sa is not None:
    # fillna with per-row mu_r values preserving index
    mu_a = sa.fillna(pd.Series(mu_r, index=sa.index)).to_numpy()
else:
    mu_a = mu_r
scale_r = np.abs(pct * mu_r)
scale_a = np.abs(pct * mu_a)
# when scale==0, normal with scale=0 returns mean
draws_r = rng.normal(loc=mu_r, scale=scale_r)
draws_a = rng.normal(loc=mu_a, scale=scale_a)
draws_r = np.clip(draws_r, 0.0, None)
draws_a = np.clip(draws_a, 0.0, None)
# Avoid zero sigmas — replace zeros with a tiny positive epsilon (arcsec)
eps = 1e-6
draws_r = np.where(draws_r <= 0.0, eps, draws_r)
draws_a = np.where(draws_a <= 0.0, eps, draws_a)
# write into per-MM columns
if 'sigma_rad [arcsec]' in df.columns:
    df['sigma_rad [arcsec]'] = draws_r
else:
    df.insert(df.shape[1], 'sigma_rad [arcsec]', draws_r)
if 'sigma_azi [arcsec]' in df.columns:
    df['sigma_azi [arcsec]'] = draws_a
else:
    df.insert(df.shape[1], 'sigma_azi [arcsec]', draws_a)
# save replacing sheet
with pd.ExcelWriter(new, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='MM_PSF', index=False)
print('wrote sampled sigmas into', new.name)
# run main.py for metrics
cmd = ['python3','main.py','-f', str(new), '--placement','elliptical','--return_metrics_only']
print('running main.py...')
proc = subprocess.run(cmd, capture_output=True, text=True)
print('returncode', proc.returncode)
print('stdout:\n', proc.stdout)
print('stderr:\n', proc.stderr)
# try parse json
try:
    metrics = json.loads(proc.stdout)
    print('METRICS_JSON:')
    print(json.dumps(metrics, indent=2))
except Exception as e:
    print('Failed parsing JSON:', e)
    print('RAW OUTPUT:\n', proc.stdout)
