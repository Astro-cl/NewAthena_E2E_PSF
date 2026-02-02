import pandas as pd
import numpy as np
import sys

path = 'Distributions/NewTest_Distribution.xlsx'
import pandas as pd
import numpy as np
import re
from pathlib import Path
import sys
# Ensure workspace root is on sys.path so we can import `main`
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import main as mainmod

path = 'Distributions/NewTest_Distribution.xlsx'
mms = [300, 100]


def norm(s):
    if s is None:
        return ''
    return re.sub(r"\s+", ' ', str(s).replace('µ', 'u').replace('μ', 'u').strip().lower())


def extract_z_from_row(r):
    for c in r.index:
        cn = norm(c)
        if 'd_align' in cn and 'z' in cn:
            v = r.get(c)
            num = pd.to_numeric(v, errors='coerce')
            if pd.notna(num):
                # if header mentions µ or value is large, assume µm
                if 'u' in str(c).lower() or abs(float(num)) > 1.0:
                    return float(num) * 1e-6
                return float(num)
    # fallback
    for c in ['d_align_z [µm]', 'd_align_z [um]', 'd_align_z', 'D_ALIGN_Z']:
        if c in r.index:
            v = r.get(c)
            num = pd.to_numeric(v, errors='coerce')
            if pd.notna(num):
                if 'µm' in c.lower() or abs(float(num)) > 1.0:
                    return float(num) * 1e-6
                return float(num)
    return 0.0


def extract_rotazi_from_row(r):
    for c in r.index:
        cn = norm(c)
        if 'rotazi' in cn or 'rot_azi' in cn or ('rot' in cn and 'azi' in cn):
            v = r.get(c)
            num = pd.to_numeric(v, errors='coerce')
            if pd.notna(num):
                return float(num)
    for c in ['d_align_rotazi', 'd_align_rotazi [arcsec]']:
        if c in r.index:
            v = r.get(c)
            num = pd.to_numeric(v, errors='coerce')
            if pd.notna(num):
                return float(num)
    return 0.0


def extract_rotrad_from_row(r):
    for c in r.index:
        cn = norm(c)
        if 'rotrad' in cn or 'rot_rad' in cn or ('rot' in cn and 'rad' in cn):
            v = r.get(c)
            num = pd.to_numeric(v, errors='coerce')
            if pd.notna(num):
                return float(num)
    for c in ['d_align_rotrad', 'd_align_rotrad [arcsec]']:
        if c in r.index:
            v = r.get(c)
            num = pd.to_numeric(v, errors='coerce')
            if pd.notna(num):
                return float(num)
    return 0.0


def interp_vig_from_rotazi(path, rotazi):
    try:
        vdf = pd.read_excel(path, sheet_name='Vignetting rotazi', engine='openpyxl')
        if vdf.shape[1] >= 2:
            xcol = pd.to_numeric(vdf.iloc[:, 0], errors='coerce')
            ycol = pd.to_numeric(vdf.iloc[:, 1], errors='coerce')
            ok = (~np.isnan(xcol)) & (~np.isnan(ycol))
            if ok.sum() > 0:
                xs = xcol[ok].to_numpy(dtype=float)
                ys = ycol[ok].to_numpy(dtype=float)
                if not np.all(np.diff(xs) > 0):
                    idx = np.argsort(xs)
                    xs = xs[idx]
                    ys = ys[idx]
                return float(np.interp(rotazi, xs, ys, left=ys[0], right=ys[-1]))
    except Exception:
        pass
    return 1.0


def interp_vig_from_rotrad(path, rotrad):
    try:
        vdf = pd.read_excel(path, sheet_name='Vignetting rotrad', engine='openpyxl')
        if vdf.shape[1] >= 2:
            xcol = pd.to_numeric(vdf.iloc[:, 0], errors='coerce')
            ycol = pd.to_numeric(vdf.iloc[:, 1], errors='coerce')
            ok = (~np.isnan(xcol)) & (~np.isnan(ycol))
            if ok.sum() > 0:
                xs = xcol[ok].to_numpy(dtype=float)
                ys = ycol[ok].to_numpy(dtype=float)
                if not np.all(np.diff(xs) > 0):
                    idx = np.argsort(xs)
                    xs = xs[idx]
                    ys = ys[idx]
                return float(np.interp(rotrad, xs, ys, left=ys[0], right=ys[-1]))
    except Exception:
        pass
    return 1.0


def build_mm_maps(path):
    mm_to_pos = {}
    mm_config_map = {}
    mm_cfg = pd.read_excel(path, sheet_name='MM configuration', engine='openpyxl')
    for i, row in mm_cfg.iterrows():
        mmn = row.get('MM #')
        if pd.isna(mmn):
            continue
        mmn = int(mmn)
        if 'Position #' in mm_cfg.columns:
            posv = row.get('Position #')
            try:
                pos = int(float(posv))
            except Exception:
                pos = i + 1
        else:
            pos = i + 1
        mm_to_pos[mmn] = pos
        x = row.get('x_MM [m]', row.get('x_MM', 0))
        y = row.get('y_MM [m]', row.get('y_MM', 0))
        z = row.get('z_MM [m]', row.get('z_MM', 0))
        r = row.get('r_MM [m]', row.get('r_MM', 0))
        mm_config_map[mmn] = {
            'x_MM': float(x or 0.0),
            'y_MM': float(y or 0.0),
            'z_MM': float(z or 0.0),
            'r_MM': float(r or 0.0),
        }
    return mm_to_pos, mm_config_map


def main():
    mm_to_pos, mm_config_map = build_mm_maps(path)
    align = pd.read_excel(path, sheet_name='Alignment', engine='openpyxl')
    # find position column
    pos_col = None
    for c in align.columns:
        s = str(c).lower()
        if 'position' in s or s.strip().startswith('pos'):
            pos_col = c
            break
    if pos_col is None:
        align['Position #'] = pd.to_numeric(align.index + 1, errors='coerce')
    else:
        align['Position #'] = pd.to_numeric(align[pos_col], errors='coerce')

    aeff_map, aeff_col = mainmod.load_aeff_weight_map_with_name(path)

    for mm in mms:
        pos = mm_to_pos.get(mm)
        print('\nMM', mm, '-> pos', pos)

        row = align[align['Position #'] == pos]
        if not row.empty:
            r = row.iloc[0]
            d_align_z = extract_z_from_row(r)
            d_align_rotazi = extract_rotazi_from_row(r)
            d_align_rotrad = extract_rotrad_from_row(r)
        else:
            d_align_z = 0.0
            d_align_rotazi = 0.0
            d_align_rotrad = 0.0

        d_grav_z = 0.0
        d_therm_z = 0.0
        d_z_total = d_align_z + d_grav_z + d_therm_z

        cfg = mm_config_map.get(mm, {'x_MM': 0.0, 'y_MM': 0.0, 'z_MM': 0.0})
        x_MM = cfg.get('x_MM', 0.0)
        y_MM = cfg.get('y_MM', 0.0)
        z_MM = cfg.get('z_MM', 0.0)
        den = 12 - z_MM
        if den != 0 and d_z_total != 0:
            dm_x = d_z_total * x_MM / den
            dm_y = d_z_total * y_MM / den
        else:
            dm_x = 0.0
            dm_y = 0.0

        vig_azi = interp_vig_from_rotazi(path, d_align_rotazi)
        vig_rad = interp_vig_from_rotrad(path, d_align_rotrad)

        aeff = aeff_map.get(mm, None)
        if aeff is None:
            aeff_str = 'N/A'
            aeff_adj_str = 'N/A'
        else:
            aeff_str = f"{aeff:.6e}"
            aeff_adj = aeff * float(vig_azi) * float(vig_rad)
            aeff_adj_str = f"{aeff_adj:.6e}"

        print(f"- d_align_z: {d_align_z:.6e} m")
        print(f"- d_grav_z:  {d_grav_z:.6e} m")
        print(f"- d_therm_z: {d_therm_z:.6e} m")
        print(f"- d_z_total: {d_z_total:.6e} m")
        print(f"- x_MM: {x_MM:.6e} m, y_MM: {y_MM:.6e} m, z_MM: {z_MM:.6e} m")
        print(f"- dm_x (m): {dm_x:.6e}")
        print(f"- dm_y (m): {dm_y:.6e}")
        print(f"- d_vignetting_rotazi (arcsec): {d_align_rotazi}")
        print(f"- d_vignetting_rotrad (arcsec): {d_align_rotrad}")
        print(f"- vig_rad factor: {vig_rad}")
        print(f"- vig_azi factor: {vig_azi}")
        print(f"- combined vig factor: {float(vig_azi) * float(vig_rad)}")
        print(f"- A_eff ({aeff_col}): {aeff_str}")
        print(f"- A_eff adjusted: {aeff_adj_str}")


if __name__ == '__main__':
    main()
