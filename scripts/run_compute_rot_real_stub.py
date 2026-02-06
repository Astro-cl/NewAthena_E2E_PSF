import time
import pandas as pd
from pathlib import Path

# Define compute_total_rot_polar copied from main.py (isolated)
def compute_total_rot_polar(mm_to_pos: dict, mm_config_map: dict, alignment_by_pos: dict, gravity_by_pos: dict, thermal_by_pos: dict):
    rotx = {}
    roty = {}
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
            rtx_total += float(gravity_by_pos[pos].get('d_grav_rotx', 0.0) or 0.0)
            rty_total += float(gravity_by_pos[pos].get('d_grav_roty', 0.0) or 0.0)
        if thermal_by_pos and pos in thermal_by_pos:
            rtx_total += float(thermal_by_pos[pos].get('d_therm_rotx', 0.0) or 0.0)
            rty_total += float(thermal_by_pos[pos].get('d_therm_roty', 0.0) or 0.0)

        rotx[pos] = rtx_total
        roty[pos] = rty_total

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
        if gravity_by_pos and pos in gravity_by_pos:
            direct_rotrad += float(gravity_by_pos[pos].get('d_grav_rotrad', 0.0) or 0.0)
            direct_rotazi += float(gravity_by_pos[pos].get('d_grav_rotazi', 0.0) or 0.0)
        if thermal_by_pos and pos in thermal_by_pos:
            direct_rotrad += float(thermal_by_pos[pos].get('d_therm_rotrad', 0.0) or 0.0)
            direct_rotazi += float(thermal_by_pos[pos].get('d_therm_rotazi', 0.0) or 0.0)

        total_rotrad = proj_rotrad + direct_rotrad
        total_rotazi = proj_rotazi + direct_rotazi

        rot_rad[pos] = total_rotrad
        rot_azi[pos] = total_rotazi

    return rotx, roty, rot_rad, rot_azi


# Load workbook and build maps like GUI does
path = Path('Distributions/Test_Distribution1.xlsx')
print('Workbook:', path)
mmc_df = pd.read_excel(path, sheet_name='MM configuration', engine='openpyxl')
print('MM rows:', len(mmc_df))

mm_to_pos = {}
mm_config_map = {}
for order_i, (_, row) in enumerate(mmc_df.iterrows()):
    mm_num = row.get('MM #')
    if pd.isna(mm_num):
        continue
    mm_num_i = int(mm_num)
    if 'Position #' in mmc_df.columns:
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
    mm_config_map[mm_num_i] = {
        'x_MM': row.get('x_MM [m]', 0),
        'y_MM': row.get('y_MM [m]', 0),
        'z_MM': row.get('z_MM [m]', 0),
        'r_MM': row.get('r_MM [m]', 0),
    }
print('Built mm_to_pos:', len(mm_to_pos))

# read perturbations

def read_pos_map(sheet_name, rot_keys):
    out = {}
    try:
        dfp = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
        if 'Position #' in dfp.columns:
            for _, r in dfp.iterrows():
                pos = r.get('Position #')
                if pd.isna(pos):
                    continue
                pos_i = int(pos)
                out[pos_i] = {k: r.get(k, 0.0) for k in rot_keys}
        elif 'MM #' in dfp.columns and mm_to_pos:
            for _, r in dfp.iterrows():
                mmn = r.get('MM #')
                if pd.isna(mmn):
                    continue
                pos = mm_to_pos.get(int(mmn))
                if pos is None:
                    continue
                out[pos] = {k: r.get(k, 0.0) for k in rot_keys}
    except Exception as e:
        print('Read', sheet_name, 'error', e)
    return out

alignment_by_pos = read_pos_map('Alignment', ['d_align_rotrad', 'd_align_rotazi'])
gravity_by_pos = read_pos_map('Gravity offload', ['d_grav_rotx', 'd_grav_roty', 'd_grav_rotrad', 'd_grav_rotazi'])
thermal_by_pos = read_pos_map('Thermal', ['d_therm_rotx', 'd_therm_roty', 'd_therm_rotrad', 'd_therm_rotazi'])

print('Calling compute_total_rot_polar...')
start = time.time()
rotx, roty, rot_rad, rot_azi = compute_total_rot_polar(mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos)
print('Returned in', time.time()-start, 's')
print('rot_rad count', len(rot_rad))
print('sample keys', sorted(list(rot_rad.keys()))[:10])
