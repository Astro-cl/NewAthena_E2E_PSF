import numpy as np
import sys, os
sys.path.insert(0, os.getcwd())
from main import load_gaussians_from_excel
p='Distributions/NewTest_Distribution.xlsx'
df = load_gaussians_from_excel(p)
if 'aeff_adjusted' in df.columns:
    weight_arr_for_center = df['aeff_adjusted'].to_numpy(dtype=float, copy=False)
elif 'weight' in df.columns:
    weight_arr_for_center = df['weight'].to_numpy(dtype=float, copy=False)
else:
    weight_arr_for_center = np.ones(len(df), dtype=float)

total_weight = float(np.nansum(weight_arr_for_center)) if weight_arr_for_center.size else 0.0
if total_weight and np.isfinite(total_weight) and total_weight>0:
    weight_arr = weight_arr_for_center / total_weight
else:
    weight_arr = np.ones(len(df), dtype=float)

for mm in (100,300):
    rows = df[df['MM #']==mm]
    if rows.empty:
        print('MM',mm,'not found')
        continue
    i = rows.index[0]
    print('MM',mm,'aeff_base',rows.iloc[0]['aeff_base'],'aeff_adjusted',rows.iloc[0]['aeff_adjusted'],'norm_amp',weight_arr[int(i)])

if (not df[df['MM #']==100].empty) and (not df[df['MM #']==300].empty):
    i100 = df[df['MM #']==100].index[0]
    i300 = df[df['MM #']==300].index[0]
    print('ratio MM300/MM100 =', weight_arr[int(i300)]/weight_arr[int(i100)])
