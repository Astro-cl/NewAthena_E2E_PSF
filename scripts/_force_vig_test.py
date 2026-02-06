import os
import glob
import numpy as np
import pandas as pd
import sys
from pathlib import Path

repo = Path(__file__).resolve().parents[1]
input_dir = repo / 'sensitivity' / 'input'
if not input_dir.exists():
    print('No sensitivity/input directory found')
    sys.exit(1)
files = sorted(list(input_dir.glob('*.xlsx')), key=lambda p: p.stat().st_mtime, reverse=True)
if not files:
    print('No generated workbooks found in sensitivity/input')
    sys.exit(1)

latest = files[0]
print('Using workbook:', latest)

# Ensure repo root is on sys.path and import main
import sys
sys.path.insert(0, str(repo))
import main

path = str(latest)

# Build MM configuration maps (mm_to_pos, pos_to_cfg_row, mm_config_map)
mm_to_pos = {}
pos_to_cfg_row = {}
mm_config_map = {}
try:
    mm_config_df = pd.read_excel(path, sheet_name='MM configuration', engine='openpyxl')
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
            cfg_row_number = int(order_i) + 1
            try:
                pos_for_row = mm_to_pos.get(mm_num_i, cfg_row_number)
            except Exception:
                pos_for_row = cfg_row_number
            pos_to_cfg_row[pos_for_row] = cfg_row_number
            mm_config_map[mm_num_i] = {
                'x_MM': row.get('x_MM [m]', 0),
                'y_MM': row.get('y_MM [m]', 0),
                'z_MM': row.get('z_MM [m]', 0),
                'r_MM': row.get('r_MM [m]', 0)
            }
except Exception as e:
    print('Could not read MM configuration:', e)

print('Found', len(mm_to_pos), 'MMs')

# Force small non-zero rotations for every position
alignment_by_pos = {}
gravity_by_pos = {}
thermal_by_pos = {}
for pos in set(mm_to_pos.values()):
    # small forced values (arcsec)
    gravity_by_pos[pos] = {'d_grav_rotx': 0.5, 'd_grav_roty': 0.2, 'd_grav_rotrad': 0.0, 'd_grav_rotazi': 0.0}
    alignment_by_pos[pos] = {'d_align_rotrad': 0.1, 'd_align_rotazi': 0.3}
    thermal_by_pos[pos] = {'d_therm_rotx': 0.0, 'd_therm_roty': 0.0}

# Compute rot projections
_, _, rot_rad_map, rot_azi_map = main.compute_total_rot_polar(mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos)

# Parse A_eff to get base mapping and selected energy
try:
    aeff_map, aeff_col_name = main.load_aeff_weight_map_with_name(path)
except Exception:
    aeff_map, aeff_col_name = main.load_aeff_weight_map(path), None
sel_energy = None
if isinstance(aeff_col_name, str):
    import re
    m = re.search(r"(\d+(?:\.\d*)?)\s*(?:keV)?", aeff_col_name, flags=re.IGNORECASE)
    if m:
        sel_energy = float(m.group(1))
if sel_energy is None:
    try:
        aeff_df = pd.read_excel(path, sheet_name='A_eff', engine='openpyxl', header=None)
        import re
        found = None
        for _, row in aeff_df.iterrows():
            for cell in row.tolist():
                try:
                    if isinstance(cell, str):
                        m = re.search(r"(\d+(?:\.\d*)?)\s*keV", cell, flags=re.IGNORECASE)
                        if m:
                            found = float(m.group(1))
                            break
                except Exception:
                    continue
            if found is not None:
                break
        if found is not None:
            sel_energy = found
    except Exception:
        sel_energy = None

print('Selected energy:', sel_energy, 'A_eff column:', aeff_col_name)

# Parse vignetting rotazi sheet to build ys_by_pos_azi
vdf_azi = pd.read_excel(path, sheet_name='Vignetting rotazi', engine='openpyxl')
ys_by_pos_azi = {}
if vdf_azi is not None and not vdf_azi.empty and vdf_azi.shape[1] >= 2:
    if vdf_azi.shape[1] >= 11:
        col_H = vdf_azi.iloc[:, 7]
        if col_H.notna().any():
            for _, r in vdf_azi.iterrows():
                try:
                    cfg_row = r.iloc[7]
                    if pd.isna(cfg_row):
                        continue
                    cfg_row = int(float(cfg_row))
                except Exception:
                    continue
                energy_marker = r.iloc[9] if vdf_azi.shape[1] > 9 else None
                try:
                    # column I is arcmin; convert to arcsec
                    xval = float(r.iloc[8]) * 60.0 if not pd.isna(r.iloc[8]) else None
                except Exception:
                    xval = None
                try:
                    yval = float(r.iloc[10]) if vdf_azi.shape[1] > 10 and not pd.isna(r.iloc[10]) else None
                except Exception:
                    yval = None
                if xval is None or yval is None:
                    continue
                key_str = (cfg_row, str(energy_marker).strip())
                key_num = None
                try:
                    key_num = (cfg_row, float(energy_marker))
                except Exception:
                    key_num = None
                key = key_str
                if key not in ys_by_pos_azi:
                    ys_by_pos_azi[key] = {'xs': [], 'ys': []}
                ys_by_pos_azi[key]['xs'].append(xval)
                ys_by_pos_azi[key]['ys'].append(yval)
                if key_num is not None:
                    if key_num not in ys_by_pos_azi:
                        ys_by_pos_azi[key_num] = {'xs': [], 'ys': []}
                    ys_by_pos_azi[key_num]['xs'].append(xval)
                    ys_by_pos_azi[key_num]['ys'].append(yval)
            for k, v in list(ys_by_pos_azi.items()):
                order = np.argsort(v['xs'])
                xs_sorted = np.array(v['xs'], dtype=float)[order]
                ys_sorted = np.array(v['ys'], dtype=float)[order]
                ys_by_pos_azi[k] = (xs_sorted, ys_sorted)

# Parse vignetting rotrad sheet to build ys_by_pos_rad
vdf_rad = pd.read_excel(path, sheet_name='Vignetting rotrad', engine='openpyxl')
ys_by_pos_rad = {}
if vdf_rad is not None and not vdf_rad.empty and vdf_rad.shape[1] >= 2:
    if vdf_rad.shape[1] >= 11:
        col_H = vdf_rad.iloc[:, 7]
        if col_H.notna().any():
            for _, r in vdf_rad.iterrows():
                try:
                    cfg_row = r.iloc[7]
                    if pd.isna(cfg_row):
                        continue
                    cfg_row = int(float(cfg_row))
                except Exception:
                    continue
                energy_marker = r.iloc[9] if vdf_rad.shape[1] > 9 else None
                try:
                    # column I is arcmin; convert to arcsec
                    xval = float(r.iloc[8]) * 60.0 if not pd.isna(r.iloc[8]) else None
                except Exception:
                    xval = None
                try:
                    yval = float(r.iloc[10]) if vdf_rad.shape[1] > 10 and not pd.isna(r.iloc[10]) else None
                except Exception:
                    yval = None
                if xval is None or yval is None:
                    continue
                key_str = (cfg_row, str(energy_marker).strip())
                key_num = None
                try:
                    key_num = (cfg_row, float(energy_marker))
                except Exception:
                    key_num = None
                key = key_str
                if key not in ys_by_pos_rad:
                    ys_by_pos_rad[key] = {'xs': [], 'ys': []}
                ys_by_pos_rad[key]['xs'].append(xval)
                ys_by_pos_rad[key]['ys'].append(yval)
                if key_num is not None:
                    if key_num not in ys_by_pos_rad:
                        ys_by_pos_rad[key_num] = {'xs': [], 'ys': []}
                    ys_by_pos_rad[key_num]['xs'].append(xval)
                    ys_by_pos_rad[key_num]['ys'].append(yval)
            for k, v in list(ys_by_pos_rad.items()):
                order = np.argsort(v['xs'])
                xs_sorted = np.array(v['xs'], dtype=float)[order]
                ys_sorted = np.array(v['ys'], dtype=float)[order]
                ys_by_pos_rad[k] = (xs_sorted, ys_sorted)

# Now compute per-MM factors and adjusted A_eff
out = []
for mm in sorted(mm_to_pos.keys())[:50]:
    pos = mm_to_pos[mm]
    cfg_row = pos_to_cfg_row.get(pos)
    # default factors
    f_azi = 1.0
    f_rad = 1.0
    # try numeric energy key first
    if cfg_row is not None and sel_energy is not None:
        keyn = (cfg_row, float(sel_energy))
        if keyn in ys_by_pos_azi:
            xs, ys = ys_by_pos_azi[keyn]
            f_azi = float(np.interp(abs(float(rot_azi_map.get(pos, 0.0))), xs, ys))
        if keyn in ys_by_pos_rad:
            xs, ys = ys_by_pos_rad[keyn]
            f_rad = float(np.interp(abs(float(rot_rad_map.get(pos, 0.0))), xs, ys))
    # fallbacks to string key or any key for cfg_row
    if cfg_row is not None and (f_azi == 1.0):
        kstr = (cfg_row, str(aeff_col_name).strip() if aeff_col_name else '')
        if kstr in ys_by_pos_azi:
            xs, ys = ys_by_pos_azi[kstr]
            f_azi = float(np.interp(abs(float(rot_azi_map.get(pos, 0.0))), xs, ys))
    if cfg_row is not None and (f_rad == 1.0):
        kstr = (cfg_row, str(aeff_col_name).strip() if aeff_col_name else '')
        if kstr in ys_by_pos_rad:
            xs, ys = ys_by_pos_rad[kstr]
            f_rad = float(np.interp(abs(float(rot_rad_map.get(pos, 0.0))), xs, ys))
    base = aeff_map.get(mm, float('nan'))
    adjusted = base * f_azi * f_rad if not np.isnan(base) else float('nan')
    out.append((mm, pos, cfg_row, base, f_azi, f_rad, adjusted))

# Print first 8 MMs
print('\nFirst 8 MM results:')
for row in out[:8]:
    mm, pos, cfg_row, base, f_azi, f_rad, adjusted = row
    print(f"MM {mm} pos {pos} cfg_row {cfg_row} base {base:.6g} f_azi {f_azi:.6g} f_rad {f_rad:.6g} adjusted {adjusted:.6g}")

print('\nDone.')
