"""Debug runner for vignetting application with stepwise prints.

Use this to identify where the apply-vignetting helper hangs.
"""
import sys
from pathlib import Path
import pandas as pd
import numpy as np

TARGET = Path('Distributions/Test_Distribution.xlsx')

def main(path=TARGET):
    print('STEP: start')
    p = Path(path)
    if not p.exists():
        print('ERROR: target not found', p)
        return 2

    print('STEP: read MM configuration')
    try:
        mmc = pd.read_excel(p, sheet_name='MM configuration', engine='openpyxl')
        print(' MM config rows:', len(mmc))
    except Exception as e:
        print(' ERROR reading MM configuration:', e)
        mmc = None

    print('STEP: build mm_to_pos and mm_config_map')
    mm_to_pos = {}
    pos_to_cfg_row = {}
    mm_config_map = {}
    if mmc is not None:
        try:
            for order_i, (_, row) in enumerate(mmc.iterrows()):
                mmnum = row.get('MM #')
                if pd.isna(mmnum):
                    continue
                mmn = int(mmnum)
                if 'Position #' in mmc.columns:
                    posv = row.get('Position #')
                    if not pd.isna(posv):
                        mm_to_pos[mmn] = int(float(posv))
                    if mmn not in mm_to_pos:
                        mm_to_pos[mmn] = order_i + 1
                else:
                    mm_to_pos[mmn] = order_i + 1
                cfg = order_i + 1
                pos = mm_to_pos.get(mmn, cfg)
                pos_to_cfg_row[pos] = cfg
                mm_config_map[mmn] = {'x_MM': row.get('x_MM [m]',0)}
            print(' mm_to_pos size', len(mm_to_pos))
        except Exception as e:
            print(' ERROR building maps', e)

    print('STEP: read perturbation sheets')
    def read_pos_map(sheet, keys):
        try:
            df = pd.read_excel(p, sheet_name=sheet, engine='openpyxl')
        except Exception as e:
            print('  could not read', sheet, e)
            return {}
        out = {}
        if 'Position #' in df.columns:
            for _, r in df.iterrows():
                pos = r.get('Position #')
                if pd.isna(pos):
                    continue
                out[int(pos)] = {k: r.get(k, 0.0) for k in keys}
        elif 'MM #' in df.columns and mm_to_pos:
            for _, r in df.iterrows():
                mm = r.get('MM #')
                if pd.isna(mm):
                    continue
                pos = mm_to_pos.get(int(mm))
                if pos is None:
                    continue
                out[pos] = {k: r.get(k, 0.0) for k in keys}
        return out

    alignment = read_pos_map('Alignment', ['d_align_rotrad', 'd_align_rotazi'])
    gravity = read_pos_map('Gravity offload', ['d_grav_rotrad','d_grav_rotazi'])
    thermal = read_pos_map('Thermal', ['d_therm_rotrad','d_therm_rotazi'])
    print(' perturbations sizes:', len(alignment), len(gravity), len(thermal))

    print('STEP: compute rot projections via main.compute_total_rot_polar')
    try:
        sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
        import main
        _, _, rot_rad, rot_azi = main.compute_total_rot_polar(mm_to_pos, mm_config_map, alignment, gravity, thermal)
        print(' rot maps sizes:', len(rot_rad), len(rot_azi))
    except Exception as e:
        print(' ERROR computing rot projections:', e)
        rot_rad = {}
        rot_azi = {}

    print('STEP: parse vignetting sheets')
    try:
        vdf_azi = pd.read_excel(p, sheet_name='Vignetting rotazi', engine='openpyxl', header=None)
        vdf_rad = pd.read_excel(p, sheet_name='Vignetting rotrad', engine='openpyxl', header=None)
        print(' vignetting shapes', vdf_azi.shape, vdf_rad.shape)
    except Exception as e:
        print(' ERROR reading vignetting sheets', e)
        vdf_azi = vdf_rad = None

    print('STEP: build per-(cfg_row,energy) arrays (quick)')
    def build(vdf):
        out = {}
        if vdf is None:
            return out
        if vdf.shape[1] < 11:
            return out
        for _, r in vdf.iterrows():
            try:
                cfg = r.iloc[7]
                if pd.isna(cfg):
                    continue
                cfg = int(float(cfg))
                x = r.iloc[8]
                e = r.iloc[9]
                y = r.iloc[10]
                if pd.isna(x) or pd.isna(y):
                    continue
                key = (cfg, e)
                out.setdefault(key, {'xs':[], 'ys':[]})
                out[key]['xs'].append(float(x))
                out[key]['ys'].append(float(y))
            except Exception:
                continue
        for k,v in list(out.items()):
            order = np.argsort(v['xs'])
            out[k] = (np.array(v['xs'])[order], np.array(v['ys'])[order])
        return out

    ys_azi = build(vdf_azi)
    ys_rad = build(vdf_rad)
    print(' built vig entries:', len(ys_azi), len(ys_rad))

    print('STEP: compute vig vals for first 10 MMs')
    count=0
    for mm,pos in list(mm_to_pos.items())[:10]:
        cfg = pos_to_cfg_row.get(pos)
        val_azi = 1.0
        if cfg is not None:
            # pick any entry
            matches = [k for k in ys_azi.keys() if k[0]==cfg]
            if matches:
                xs,ys = ys_azi[matches[0]]
                xv = float(rot_azi.get(pos,0.0))
                try:
                    val_azi = float(np.interp(xv,xs,ys))
                except Exception as e:
                    print(' interp error', e)
        val_rad = 1.0
        if cfg is not None:
            matches = [k for k in ys_rad.keys() if k[0]==cfg]
            if matches:
                xs,ys = ys_rad[matches[0]]
                xv = float(rot_rad.get(pos,0.0))
                try:
                    val_rad = float(np.interp(xv,xs,ys))
                except Exception as e:
                    print(' interp rad error', e)
        print(' MM', mm, 'pos', pos, 'cfg', cfg, 'azi', val_azi, 'rad', val_rad)
        count+=1

    print('STEP: done')
    return 0

if __name__ == '__main__':
    sys.exit(main())
