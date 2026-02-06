import pandas as pd
path='Distributions/TestDistribution6_working.xlsx'
# Build mm_to_pos mapping as main does
mm_to_pos={}
mmc=pd.read_excel(path,sheet_name='MM configuration',engine='openpyxl')
for i,row in mmc.iterrows():
    mm_num=row.get('MM #')
    if pd.isna(mm_num):
        continue
    mm_to_pos[int(mm_num)]=int(i)+1
# Build mm_config_map
mm_config_map={}
for _,row in mmc.iterrows():
    mm_num=row.get('MM #')
    if pd.isna(mm_num):
        continue
    mm_config_map[int(mm_num)]={'x_MM':row.get('x_MM [m]',0),'y_MM':row.get('y_MM [m]',0),'r_MM':row.get('r_MM [m]',0),'z_MM':row.get('z_MM [m]',0)}
# simple loaders for alignment/gravity/thermal
def load_sheet_simple(name):
    try:
        df=pd.read_excel(path,sheet_name=name,engine='openpyxl')
    except Exception:
        return {}
    d={}
    if 'Position #' in df.columns:
        tmp=df.copy()
        tmp['Position #']=pd.to_numeric(tmp['Position #'],errors='coerce')
        tmp=tmp[tmp['Position #'].notna()]
        for _,r in tmp.iterrows():
            pos=int(r['Position #'])
            d[pos]={}
            d[pos]['d_align_rotrad']=float(r.get('d_align_rotrad [arcsec]',0) or 0)
            d[pos]['d_align_rotazi']=float(r.get('d_align_rotazi [arcsec]',0) or 0)
    return d
alignment=load_sheet_simple('Alignment')
gravity=load_sheet_simple('Gravity offload')
thermal=load_sheet_simple('Thermal')

# Reimplement minimal compute_total_rot_polar logic focusing on rot_azi/rot_rad
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

    pos_to_mm = {}
    if mm_to_pos:
        for mm, p in mm_to_pos.items():
            pos_to_mm.setdefault(p, mm)

    for pos in positions:
        rtx_total = 0.0
        rty_total = 0.0
        if gravity_by_pos and pos in gravity_by_pos:
            rtx_total += float(gravity_by_pos[pos].get('d_align_rotrad', 0.0) or 0.0)
            rty_total += float(gravity_by_pos[pos].get('d_align_rotazi', 0.0) or 0.0)
        if thermal_by_pos and pos in thermal_by_pos:
            rtx_total += float(thermal_by_pos[pos].get('d_align_rotrad', 0.0) or 0.0)
            rty_total += float(thermal_by_pos[pos].get('d_align_rotazi', 0.0) or 0.0)

        ux, uy = 1.0, 0.0
        mm_choice = pos_to_mm.get(pos)
        if mm_choice is not None and mm_choice in mm_config_map:
            cfg = mm_config_map.get(mm_choice, {})
            r_mm = float(cfg.get('r_MM', 0.0) or 0.0)
            x_mm = float(cfg.get('x_MM', 0.0) or 0.0)
            y_mm = float(cfg.get('y_MM', 0.0) or 0.0)
            if r_mm > 0.0:
                ux = x_mm / r_mm
                uy = y_mm / r_mm

        proj_rotrad = rtx_total * ux + rty_total * uy
        proj_rotazi = -rtx_total * uy + rty_total * ux

        direct_rotrad = 0.0
        direct_rotazi = 0.0
        if alignment_by_pos and pos in alignment_by_pos:
            direct_rotrad += float(alignment_by_pos[pos].get('d_align_rotrad', 0.0) or 0.0)
            direct_rotazi += float(alignment_by_pos[pos].get('d_align_rotazi', 0.0) or 0.0)

        total_rotrad = proj_rotrad + direct_rotrad
        total_rotazi = proj_rotazi + direct_rotazi

        rot_rad[pos] = total_rotrad
        rot_azi[pos] = total_rotazi

    return rot_rad, rot_azi

rot_rad, rot_azi = compute_rot_polar_local(mm_to_pos, mm_config_map, alignment, gravity, thermal)
print('rot_azi for pos1:', rot_azi.get(1))
print('rot_rad for pos1:', rot_rad.get(1))
