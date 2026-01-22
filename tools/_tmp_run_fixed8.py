from pathlib import Path
import tempfile, json, subprocess, sys
import pandas as pd
from numpy.random import default_rng
from tools.run_sensitivity import apply_mm_psf_choice_to_df, build_aeff_mapping, load_standard_mm_psf_presets

BASE=Path('Distributions/Test_Distribution.xlsx')
AE=build_aeff_mapping()
SP=load_standard_mm_psf_presets(BASE)
df=pd.read_excel(BASE, sheet_name='MM_PSF', engine='openpyxl')
rng=default_rng(0)
df2=apply_mm_psf_choice_to_df(df.copy(), 'Fixed Sym Gaussian 8"', AE, rng, standard_presets=SP)
# write temp workbook
with tempfile.NamedTemporaryFile(prefix='sens_', suffix='.xlsx', delete=False) as tf:
    tmpname = tf.name
with pd.ExcelWriter(tmpname, engine='openpyxl') as w:
    df2.to_excel(w, sheet_name='MM_PSF', index=False)

cmd=[sys.executable, 'main.py', '-f', tmpname, '--return_metrics_only', '--metrics-nr-final', '300', '--metrics-ntheta-final', '24', '--metrics-r-margin', '6.0']
print('CMD:', ' '.join(cmd))
proc=subprocess.run(cmd, capture_output=True, text=True, timeout=20)
print('RC:', proc.returncode)
print('--- STDOUT ---')
print(proc.stdout)
print('--- STDERR ---')
print(proc.stderr)

# try to parse stdout as JSON
try:
    print('Parsed JSON:', json.loads(proc.stdout))
except Exception as e:
    print('JSON parse failed:', e)
