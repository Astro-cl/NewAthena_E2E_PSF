import pandas as pd
import sys
sys.path.insert(0, '.')
import main
from pathlib import Path
p=Path('sensitivity/input/20260205T161117Z_1_A_eff1_keV_MM_PSFFixed_Sym_Gaussian_0.01_Alignment0_Thermal0_Gravity_offload0.xlsx')
# build mm_to_pos and mm_config_map
mm_to_pos={}
mm_config_map={}
mm_cfg=pd.read_excel(p,sheet_name='MM configuration',engine='openpyxl')
for order_i,(idx,row) in enumerate(mm_cfg.iterrows()):
    mm_num=row.get('MM #')
    if pd.isna(mm_num):
        continue
    mm_num_i=int(mm_num)
    if 'Position #' in mm_cfg.columns:
        pos_val=row.get('Position #')
        if not pd.isna(pos_val):
            try:
                mm_to_pos[mm_num_i]=int(float(pos_val))
            except:
                pass
        if mm_num_i not in mm_to_pos:
            mm_to_pos[mm_num_i]=int(order_i)+1
    else:
        mm_to_pos[mm_num_i]=int(order_i)+1
    mm_config_map[mm_num_i]={'x_MM':row.get('x_MM [m]',0),'y_MM':row.get('y_MM [m]',0),'z_MM':row.get('z_MM [m]',0),'r_MM':row.get('r_MM [m]',0)}

# load alignment/gravity/thermal by position
alignment_by_pos={}
gravity_by_pos={}
thermal_by_pos={}
for sheet, target in [('Alignment', alignment_by_pos), ('Gravity offload', gravity_by_pos), ('Thermal', thermal_by_pos)]:
    try:
        df=pd.read_excel(p,sheet_name=sheet,engine='openpyxl')
    except Exception:
        continue
    if 'Position #' in df.columns:
        tmp=df.copy()
        tmp['Position #']=pd.to_numeric(tmp['Position #'],errors='coerce')
        tmp=tmp[tmp['Position #'].notna()]
        grp=tmp.groupby('Position #',as_index=False).sum(numeric_only=True)
        for _,row in grp.iterrows():
            pos=int(row['Position #'])
            if sheet=='Alignment':
                target[pos]={'d_align_rad':float(row.get('d_align_rad [µm]',0))*1e-6,'d_align_azi':float(row.get('d_align_azi [µm]',0))*1e-6,'d_align_z':float(row.get('d_align_z [µm]',0))*1e-6,'d_align_rotz':float(row.get('d_align_rotz [arcsec]',0))}
            elif sheet=='Gravity offload':
                target[pos]={'d_grav_x':float(row.get('d_grav_x [µm]',0))*1e-6,'d_grav_y':float(row.get('d_grav_y [µm]',0))*1e-6,'d_grav_z':float(row.get('d_grav_z [µm]',0))*1e-6,'d_grav_rotz':float(row.get('d_grav_rotz [arcsec]',0))}
            else:
                target[pos]={'d_therm_x':float(row.get('d_therm_x [µm]',0))*1e-6,'d_therm_y':float(row.get('d_therm_y [µm]',0))*1e-6,'d_therm_z':float(row.get('d_therm_z [µm]',0))*1e-6,'d_therm_rotz':float(row.get('d_therm_rotz [arcsec]',0))}

# compute rot maps
rotx,roty,rot_rad_map,rot_azi_map=main.compute_total_rot_polar(mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos)
print('rot_azi_map sample:', {k:rot_azi_map.get(k) for k in sorted(list(rot_azi_map.keys()))[:10]})
print('rot_rad_map sample:', {k:rot_rad_map.get(k) for k in sorted(list(rot_rad_map.keys()))[:10]})

# read vdf_azi and find per-row-energy series for cfg_row matching pos 1
vdf_azi=pd.read_excel(p,sheet_name='Vignetting rotazi',engine='openpyxl',header=None)
# col H index 7 has cfg_row, col I index8 x, colJ index9 energy, colK index10 factor
pairs=[]
for _,r in vdf_azi.iterrows():
    if not pd.isna(r.iloc[7]):
        try:
            cfg=int(float(r.iloc[7]))
        except:
            continue
        x=r.iloc[8]
        e=r.iloc[9]
        y=r.iloc[10]
        if pd.isna(x) or pd.isna(y):
            continue
        pairs.append((cfg,e,float(x),float(y)))
# show first matches for cfg_row of position 1: need pos_to_cfg_row mapping: replicate using mm_cfg
pos_to_cfg={}
for order_i,(idx,row) in enumerate(mm_cfg.iterrows()):
    cfg_row=order_i+1
    pos_val=row.get('Position #')
    if not pd.isna(pos_val):
        try:
            pos_to_cfg[int(float(pos_val))]=cfg_row
        except:
            pos_to_cfg[cfg_row]=cfg_row
    else:
        pos_to_cfg[cfg_row]=cfg_row
print('pos_to_cfg sample:', {k:pos_to_cfg[k] for k in list(pos_to_cfg)[:8]})
cfg_for_p1=pos_to_cfg.get(1)
print('cfg_for_p1=',cfg_for_p1)
# find pairs with cfg==cfg_for_p1
p1_pairs=[(x,y) for (cfg,e,x,y) in pairs if cfg==cfg_for_p1]
print('Found p1 pairs (first 10):',p1_pairs[:10])
# interpolate at rot_azi_map[p]
import numpy as np
if p1_pairs and 1 in rot_azi_map:
    xs=np.array([xx for xx,yy in p1_pairs])
    ys=np.array([yy for xx,yy in p1_pairs])
    order=np.argsort(xs)
    xs=xs[order]; ys=ys[order]
    val=float(rot_azi_map.get(1,0.0))
    fac=np.interp(val,xs,ys)
    print('xs[:10]=',xs[:10])
    print('ys[:10]=',ys[:10])
    print('interp at',val,'->',fac)
else:
    print('no pairs or no rot_azi_map value')
