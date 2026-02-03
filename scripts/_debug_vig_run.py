import pandas as pd
import os
from main import load_gaussians_from_excel, compute_total_rot_polar

path = 'tmp_vig_test.xlsx'
mm_psf = pd.DataFrame({'MM #':[1],'m_rad [arcsec]':[0.0],'m_azi [arcsec]':[0.0],'sigma_rad [arcsec]':[4.3],'sigma_azi [arcsec]':[4.3]})
aeff = pd.DataFrame({'MM #':[1],'0.25 keV':[1.0]})
mm_conf = pd.DataFrame({'MM #':[1],'Position #':[1],'x_MM [m]':[1.0],'y_MM [m]':[0.0],'r_MM [m]':[1.0]})
grav = pd.DataFrame({'Position #':[1],'d_grav_rotx [arcsec]':[1.0],'d_grav_roty [arcsec]':[0.0]})
vig_azi = pd.DataFrame({'delta_arcsec':[-1.0,0.0,1.0],'1':[0.9,1.0,1.1]})
vig_rad = pd.DataFrame({'delta_arcsec':[-1.0,0.0,1.0],'1':[0.95,1.0,1.1]})
with pd.ExcelWriter(path, engine='openpyxl') as w:
    mm_psf.to_excel(w, sheet_name='MM_PSF', index=False)
    aeff.to_excel(w, sheet_name='A_eff', index=False)
    mm_conf.to_excel(w, sheet_name='MM configuration', index=False)
    grav.to_excel(w, sheet_name='Gravity offload', index=False)
    vig_azi.to_excel(w, sheet_name='Vignetting rotazi', index=False)
    vig_rad.to_excel(w, sheet_name='Vignetting rotrad', index=False)

print('Calling loader...')
df = load_gaussians_from_excel(path)
print('df.attrs:', df.attrs)
print('df rows:')
print(df[['MM #','weight']])
print('weights:', df['weight'].tolist())

# compute projections separately
mm_to_pos = {1:1}
mm_config_map = {1:{'x_MM':1.0,'y_MM':0.0,'r_MM':1.0}}
alignment_by_pos = {}
gravity_by_pos = {1: {'d_grav_rotx':1.0,'d_grav_roty':0.0}}
thermal_by_pos = {}
rotx, roty, rot_rad, rot_azi = compute_total_rot_polar(mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos)
print('rot_rad:', rot_rad)
print('rot_azi:', rot_azi)

# inspect sheets
vdf_azi = pd.read_excel(path, sheet_name='Vignetting rotazi', engine='openpyxl')
print('Vig azi cols:', list(vdf_azi.columns))
print(vdf_azi)
vdf_rad = pd.read_excel(path, sheet_name='Vignetting rotrad', engine='openpyxl')
print('Vig rad cols:', list(vdf_rad.columns))
print(vdf_rad)

# Reconstruct mm_to_pos as loader would
mm_config_df = pd.read_excel(path, sheet_name='MM configuration', engine='openpyxl')
mm_to_pos = {}
if 'MM #' in mm_config_df.columns:
    for order_i, (_, row) in enumerate(mm_config_df.iterrows()):
        mm_num = row.get('MM #')
        if pd.isna(mm_num):
            continue
        mm_num_i = int(mm_num)
        if 'Position #' in mm_config_df.columns:
            pos_val = row.get('Position #')
            if not pd.isna(pos_val):
                try:
                    mm_to_pos[mm_num_i] = int(float(pos_val))
                except Exception:
                    pass
            if mm_num_i not in mm_to_pos:
                mm_to_pos[mm_num_i] = int(order_i) + 1
        else:
            mm_to_pos[mm_num_i] = int(order_i) + 1

print('reconstructed mm_to_pos:', mm_to_pos)
pos = 1
mask = df['MM #'].map(lambda m: mm_to_pos.get(int(m)) if pd.notna(m) else None) == pos
print('mask values:', mask.tolist())

print('Applying factor to df using same mask...')
df.loc[mask, 'weight'] = df.loc[mask, 'weight'].astype(float) * 1.1
print(df[['MM #','weight']])

os.remove(path)
