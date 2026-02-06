import time
import pandas as pd
from pathlib import Path
import sys
from pathlib import Path as _Path
# Ensure project root is on sys.path when running from scripts/
sys.path.insert(0, str(_Path(__file__).resolve().parents[1]))
import main

path = Path('Distributions/Test_Distribution.xlsx')
print('Workbook:', path)
start = time.time()
mmc_df = pd.read_excel(path, sheet_name='MM configuration', engine='openpyxl')
print('Read MM configuration in', time.time()-start, 's; rows=', len(mmc_df))

mm_to_pos = {}
pos_to_cfg_row = {}
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
    cfg_row_number = int(order_i) + 1
    pos_for_row = mm_to_pos.get(mm_num_i, cfg_row_number)
    pos_to_cfg_row[pos_for_row] = cfg_row_number
    mm_config_map[mm_num_i] = {
        'x_MM': row.get('x_MM [m]', 0),
        'y_MM': row.get('y_MM [m]', 0),
        'z_MM': row.get('z_MM [m]', 0),
        'r_MM': row.get('r_MM [m]', 0),
    }
print('Built mm maps; mm_to_pos=', len(mm_to_pos), 'mm_config_map=', len(mm_config_map))


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

start = time.time()
alignment_by_pos = read_pos_map('Alignment', ['d_align_rotrad', 'd_align_rotazi'])
print('Read Alignment in', time.time()-start, 's; entries=', len(alignment_by_pos))
start = time.time()
gravity_by_pos = read_pos_map('Gravity offload', ['d_grav_rotx', 'd_grav_roty', 'd_grav_rotrad', 'd_grav_rotazi'])
print('Read Gravity in', time.time()-start, 's; entries=', len(gravity_by_pos))
start = time.time()
thermal_by_pos = read_pos_map('Thermal', ['d_therm_rotx', 'd_therm_roty', 'd_therm_rotrad', 'd_therm_rotazi'])
print('Read Thermal in', time.time()-start, 's; entries=', len(thermal_by_pos))

print('Calling compute_total_rot_polar...')
start = time.time()
rotx, roty, rot_rad, rot_azi = main.compute_total_rot_polar(mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos)
print('compute_total_rot_polar returned in', time.time()-start, 's')
print('rot_rad size', len(rot_rad), 'rot_azi size', len(rot_azi))

# print a few entries
keys = sorted(rot_rad.keys())[:10]
for k in keys:
    print('pos', k, 'rot_rad', rot_rad[k], 'rot_azi', rot_azi.get(k))
