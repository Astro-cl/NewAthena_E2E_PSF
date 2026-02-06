import pandas as pd
import numpy as np
path='Distributions/TestDistribution6_working.xlsx'
# mm config
mmc=pd.read_excel(path,sheet_name='MM configuration',engine='openpyxl')
mm_to_pos={}
pos_to_cfg_row={}
mm_config_map={}
for order_i,(idx,row) in enumerate(mmc.iterrows()):
    mm_num=row.get('MM #')
    if pd.isna(mm_num):
        continue
    mm_num_i=int(mm_num)
    if 'Position #' in mmc.columns:
        pos_val=row.get('Position #')
        if not pd.isna(pos_val):
            try:
                p=int(float(pos_val))
            except Exception:
                p=order_i+1
        else:
            p=order_i+1
    else:
        p=order_i+1
    mm_to_pos[mm_num_i]=p
    cfg_row_number=order_i+1
    pos_to_cfg_row[p]=cfg_row_number
    mm_config_map[mm_num_i]={'x_MM':row.get('x_MM [m]',0),'y_MM':row.get('y_MM [m]',0),'r_MM':row.get('r_MM [m]',0),'z_MM':row.get('z_MM [m]',0)}

# simple compute_total_rot_polar local
def compute_rot_polar_local(mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos):
    rot_rad = {}
    rot_azi = {}
    positions = set()
    if mm_to_pos:
        positions.update(mm_to_pos.values())
    if alignment_by_pos:
        positions.update(alignment_by_pos.keys())
    if gravity_by_pos:
        positions.update(gravity_by_pos.keys())
    if thermal_by_pos:
        positions.update(thermal_by_pos.keys())
    pos_to_mm={}
    if mm_to_pos:
        for mm,p in mm_to_pos.items():
            pos_to_mm.setdefault(p, mm)
    for pos in positions:
        rtx_total = 0.0
        rty_total = 0.0
        if gravity_by_pos and pos in gravity_by_pos:
            rtx_total += float(gravity_by_pos[pos].get('d_grav_rotx', 0.0) or 0.0)
            rty_total += float(gravity_by_pos[pos].get('d_grav_roty', 0.0) or 0.0)
        if thermal_by_pos and pos in thermal_by_pos:
            rtx_total += float(thermal_by_pos[pos].get('d_therm_rotx', 0.0) or 0.0)
            rty_total += float(thermal_by_pos[pos].get('d_therm_roty', 0.0) or 0.0)
        ux,uy=1.0,0.0
        mm_choice = pos_to_mm.get(pos)
        if mm_choice is not None and mm_choice in mm_config_map:
            cfg=mm_config_map.get(mm_choice,{})
            r_mm=float(cfg.get('r_MM',0.0) or 0.0)
            x_mm=float(cfg.get('x_MM',0.0) or 0.0)
            y_mm=float(cfg.get('y_MM',0.0) or 0.0)
            if r_mm>0.0:
                ux=x_mm/r_mm; uy=y_mm/r_mm
        proj_rotrad = rtx_total * ux + rty_total * uy
        proj_rotazi = -rtx_total * uy + rty_total * ux
        direct_rotrad=0.0; direct_rotazi=0.0
        if alignment_by_pos and pos in alignment_by_pos:
            direct_rotrad += float(alignment_by_pos[pos].get('d_align_rotrad', 0.0) or 0.0)
            direct_rotazi += float(alignment_by_pos[pos].get('d_align_rotazi', 0.0) or 0.0)
        total_rotrad = proj_rotrad + direct_rotrad
        total_rotazi = proj_rotazi + direct_rotazi
        rot_rad[pos]=total_rotrad
        rot_azi[pos]=total_rotazi
    return rot_rad, rot_azi

# load alignment/gravity/thermal by Position #
def load_pos_map(name, rot_pref=None):
    try:
        df=pd.read_excel(path,sheet_name=name,engine='openpyxl')
    except Exception:
        return {}
    d={}
    if 'Position #' in df.columns:
        tmp=df.copy(); tmp['Position #']=pd.to_numeric(tmp['Position #'],errors='coerce'); tmp=tmp[tmp['Position #'].notna()]
        for _,r in tmp.iterrows():
            pos=int(r['Position #'])
            d[pos]={}
            # include rotx/roty/rotz where present
            for k in r.index:
                kval = r.get(k, 0)
                try:
                    kvalf = float(kval)
                except Exception:
                    kvalf = 0.0
                if 'rotx' in str(k).lower(): d[pos]['d_grav_rotx' if name=='Gravity offload' else ('d_therm_rotx' if name=='Thermal' else 'd_align_rotx')] = kvalf
                if 'roty' in str(k).lower(): d[pos]['d_grav_roty' if name=='Gravity offload' else ('d_therm_roty' if name=='Thermal' else 'd_align_roty')] = kvalf
                if 'rotz' in str(k).lower(): d[pos]['d_grav_rotz' if name=='Gravity offload' else ('d_therm_rotz' if name=='Thermal' else 'd_align_rotz')] = kvalf
            # also include direct polar if present
            if 'd_align_rotrad' in r.index: d[pos]['d_align_rotrad']=float(r.get('d_align_rotrad',0) or 0)
            if 'd_align_rotazi' in r.index: d[pos]['d_align_rotazi']=float(r.get('d_align_rotazi',0) or 0)
    return d

alignment = load_pos_map('Alignment')
gravity = load_pos_map('Gravity offload')
thermal = load_pos_map('Thermal')
rot_rad_map, rot_azi_map = compute_rot_polar_local(mm_to_pos, mm_config_map, alignment, gravity, thermal)

# Parse vignetting sheets into per-(cfg_row,energy) series
vdf_azi = pd.read_excel(path, sheet_name='Vignetting rotazi', engine='openpyxl', header=None)
ys_by_pos_azi = {}
xs_azi = ys_azi = None
azi_mode='none'
if vdf_azi is not None and not vdf_azi.empty and vdf_azi.shape[1] >= 2:
    if vdf_azi.shape[1] >= 11:
        col_H = vdf_azi.iloc[:,7]
        if col_H.notna().any():
            for _, r in vdf_azi.iterrows():
                try:
                    cfg_row = r.iloc[7]
                    if pd.isna(cfg_row): continue
                    cfg_row=int(float(cfg_row))
                except Exception:
                    continue
                energy_marker = r.iloc[9] if vdf_azi.shape[1] > 9 else None
                try:
                    xval = float(r.iloc[8]) * 60.0 if not pd.isna(r.iloc[8]) else None
                except Exception:
                    xval=None
                try:
                    yval = float(r.iloc[10]) if vdf_azi.shape[1] > 10 and not pd.isna(r.iloc[10]) else None
                except Exception:
                    yval=None
                if xval is None or yval is None: continue
                key_str=(cfg_row, str(energy_marker).strip())
                key_num=None
                try:
                    key_num=(cfg_row, float(energy_marker))
                except Exception:
                    key_num=None
                for key in (key_str, key_num):
                    if key is None: continue
                    if key not in ys_by_pos_azi: ys_by_pos_azi[key]={'xs':[],'ys':[]}
                    ys_by_pos_azi[key]['xs'].append(xval)
                    ys_by_pos_azi[key]['ys'].append(yval)
            for k,v in list(ys_by_pos_azi.items()):
                order=np.argsort(v['xs'])
                xs_sorted=np.array(v['xs'],dtype=float)[order]
                ys_sorted=np.array(v['ys'],dtype=float)[order]
                ys_by_pos_azi[k]=(xs_sorted, ys_sorted)
            if ys_by_pos_azi:
                azi_mode='per_row_energy'
# detect selected energy from vdf C2
sel_energy=None
try:
    if vdf_azi.shape[0]>1 and vdf_azi.shape[1]>2:
        cand=vdf_azi.iat[1,2]
        if cand is not None:
            try: sel_energy=float(str(cand).strip())
            except Exception:
                import re
                m=re.search(r"(\d+(?:\.\d*)?)", str(cand))
                if m: sel_energy=float(m.group(1))
except Exception:
    sel_energy=None

# build vig maps per position
vig_vals_azi={}
for pos in sorted(set(mm_to_pos.values())):
    cfg_row=pos_to_cfg_row.get(pos)
    used=False
    factor=1.0
    if azi_mode=='per_row_energy' and cfg_row is not None:
        if sel_energy is not None:
            keyn=(cfg_row, float(sel_energy))
            if keyn in ys_by_pos_azi:
                xs_use, ys_use = ys_by_pos_azi[keyn]
                factor=float(np.interp(abs(float(rot_azi_map.get(pos,0.0))), xs_use, ys_use))
                used=True
        if not used:
            key=(cfg_row, str(sel_energy))
            if key in ys_by_pos_azi:
                xs_use, ys_use=ys_by_pos_azi[key]
                factor=float(np.interp(abs(float(rot_azi_map.get(pos,0.0))), xs_use, ys_use))
                used=True
        if not used:
            matches=[k for k in ys_by_pos_azi.keys() if k[0]==cfg_row]
            if matches:
                k=matches[0]
                xs_use, ys_use = ys_by_pos_azi[k]
                factor=float(np.interp(abs(float(rot_azi_map.get(pos,0.0))), xs_use, ys_use))
                used=True
    vig_vals_azi[pos]=factor

# Now for rotrad (simpler: per_row_energy same logic)
vdf_rad = pd.read_excel(path, sheet_name='Vignetting rotrad', engine='openpyxl', header=None)
ys_by_pos_rad={}
xs_rad=ys_rad=None
rad_mode='none'
if vdf_rad is not None and not vdf_rad.empty and vdf_rad.shape[1]>=2:
    if vdf_rad.shape[1]>=11:
        col_H = vdf_rad.iloc[:,7]
        if col_H.notna().any():
            for _, r in vdf_rad.iterrows():
                try:
                    cfg_row = r.iloc[7]
                    if pd.isna(cfg_row): continue
                    cfg_row=int(float(cfg_row))
                except Exception:
                    continue
                energy_marker = r.iloc[9] if vdf_rad.shape[1] > 9 else None
                try:
                    xval = float(r.iloc[8]) * 60.0 if not pd.isna(r.iloc[8]) else None
                except Exception:
                    xval=None
                try:
                    yval = float(r.iloc[10]) if vdf_rad.shape[1] > 10 and not pd.isna(r.iloc[10]) else None
                except Exception:
                    yval=None
                if xval is None or yval is None: continue
                key_str=(cfg_row, str(energy_marker).strip())
                key_num=None
                try:
                    key_num=(cfg_row, float(energy_marker))
                except Exception:
                    key_num=None
                for key in (key_str, key_num):
                    if key is None: continue
                    if key not in ys_by_pos_rad: ys_by_pos_rad[key]={'xs':[],'ys':[]}
                    ys_by_pos_rad[key]['xs'].append(xval)
                    ys_by_pos_rad[key]['ys'].append(yval)
            for k,v in list(ys_by_pos_rad.items()):
                order=np.argsort(v['xs'])
                xs_sorted=np.array(v['xs'],dtype=float)[order]
                ys_sorted=np.array(v['ys'],dtype=float)[order]
                ys_by_pos_rad[k]=(xs_sorted, ys_sorted)
            if ys_by_pos_rad:
                rad_mode='per_row_energy'

vig_vals_rad={}
for pos in sorted(set(mm_to_pos.values())):
    cfg_row=pos_to_cfg_row.get(pos)
    used=False
    factor=1.0
    if rad_mode=='per_row_energy' and cfg_row is not None:
        if sel_energy is not None:
            keyn=(cfg_row, float(sel_energy))
            if keyn in ys_by_pos_rad:
                xs_use, ys_use = ys_by_pos_rad[keyn]
                factor=float(np.interp(abs(float(rot_rad_map.get(pos,0.0))), xs_use, ys_use))
                used=True
        if not used:
            key=(cfg_row, str(sel_energy))
            if key in ys_by_pos_rad:
                xs_use, ys_use = ys_by_pos_rad[key]
                factor=float(np.interp(abs(float(rot_rad_map.get(pos,0.0))), xs_use, ys_use))
                used=True
        if not used:
            matches=[k for k in ys_by_pos_rad.keys() if k[0]==cfg_row]
            if matches:
                k=matches[0]
                xs_use, ys_use=ys_by_pos_rad[k]
                factor=float(np.interp(abs(float(rot_rad_map.get(pos,0.0))), xs_use, ys_use))
                used=True
    vig_vals_rad[pos]=factor

print('sel_energy', sel_energy)
print('rot_azi pos1', rot_azi_map.get(1))
print('rot_rad pos1', rot_rad_map.get(1))
print('vig_vals_azi pos1', vig_vals_azi.get(1))
print('vig_vals_rad pos1', vig_vals_rad.get(1))
# print a few sample positions
for p in range(1,9):
    print(p, 'azi', vig_vals_azi.get(p), 'rad', vig_vals_rad.get(p))
