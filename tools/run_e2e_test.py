import pandas as pd, tempfile, subprocess
from pathlib import Path
mm = [1,2,3,4]
mp = pd.DataFrame({
    'm_rad [arcsec]':[0.0,0.0,0.0,0.0],
    'm_azi [arcsec]':[0.0,0.0,0.0,0.0],
    'sigma_rad [arcsec]':[8.0,8.5,7.8,9.0],
    'sigma_azi [arcsec]':[3.0,2.8,3.5,2.9],
    'distribution':['gaussian']*4,
    'alpha_rad':[0.5]*4,
    'alpha_azi':[0.5]*4,
    'MM #':mm
})
mc = pd.DataFrame({'MM #':mm, 'x_MM [m]':[0.01,-0.01,0.02,-0.02], 'y_MM [m]':[1.0,1.0,1.0,1.0]})
mc['r_MM [m]'] = (mc['x_MM [m]']**2 + mc['y_MM [m]']**2)**0.5
mc['Row #'] = [1,1,2,2]
ae = pd.DataFrame({'MM #':mm, 'weight':[1.0,1.0,1.0,1.0]})
align = pd.DataFrame([[0.0,0.0,0.0]]*len(mm))
# Write multi-sheet CSV
tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.csv')
path = Path(tmp.name)
with open(path, 'w', encoding='utf-8') as fh:
    fh.write('# sheet: MM_PSF\n')
    fh.write(mp.to_csv(index=False))
    fh.write('\n# sheet: MM configuration\n')
    fh.write(mc.to_csv(index=False))
    fh.write('\n# sheet: A_eff\n')
    fh.write(ae.to_csv(index=False))
    fh.write('\n# sheet: Alignment\n')
    fh.write(align.to_csv(index=False, header=False))
print('Wrote test CSV to', path)
# Run main.py
cmd = ['python3','main.py','-f',str(path),'--input-csv',str(path),'--optimize','--placement','elliptical','--mode','coarse','--suppress-output','--return_metrics_only']
print('Running:', ' '.join(cmd))
proc = subprocess.run(cmd, text=True, capture_output=True)
print('RC', proc.returncode)
print('STDOUT:\n', proc.stdout)
print('STDERR:\n', proc.stderr)
