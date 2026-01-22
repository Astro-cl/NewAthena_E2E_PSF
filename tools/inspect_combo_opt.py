from pathlib import Path
import sys, time, hashlib, itertools, json
import pandas as pd
import numpy as np

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from tools.run_sensitivity import SENS_PATH, BASE_WORKBOOK, build_aeff_mapping, load_standard_mm_psf_presets, load_mm_row_map, find_aeff_weights_for_choice, apply_mm_psf_choice_to_df
from main import load_gaussians_from_excel, plot_sum
from optimize_mm_rows import load_all_sheets, _load_base_params_from_workbook, _load_position_deltas, _elliptical_place_mm_config, rebuild_df


def inspect_combo(idx_target=2):
    param_options = {}
    sens = pd.read_excel(SENS_PATH, sheet_name=0, engine='openpyxl', header=0)
    for col in sens.columns:
        vals = sens[col].dropna().astype(str).map(str.strip)
        vals = [v for v in vals if v not in ['', '-', 'NaN', 'nan', 'None']]
        if vals:
            param_options[col] = vals
    keys = list(param_options.keys())
    combos = list(itertools.product(*(param_options[k] for k in keys)))
    combo = combos[idx_target-1]
    combo_dict = dict(zip(keys, combo))
    print('Inspecting combo idx', idx_target, combo_dict)

    aeff_map = build_aeff_mapping()
    aeff_map['mm_row_map'] = load_mm_row_map(BASE_WORKBOOK)
    standard = load_standard_mm_psf_presets(BASE_WORKBOOK)

    h = hashlib.sha1(repr(combo_dict).encode('utf8')).hexdigest()
    seed = int(h[:16], 16) % (2**32)
    rng = np.random.default_rng(seed)

    df_run = load_gaussians_from_excel(str(BASE_WORKBOOK), sheet='MM_PSF')
    mapping = None
    if 'A_eff' in combo_dict:
        mapping = find_aeff_weights_for_choice(combo_dict['A_eff'], aeff_map)
        df_run['weight'] = df_run['MM #'].astype(int).map(mapping)
    if 'MM_PSF' in combo_dict:
        df_run = apply_mm_psf_choice_to_df(df_run, combo_dict['MM_PSF'], aeff_map, rng, standard_presets=standard)

    print('\n-- df_run summary --')
    print('rows:', len(df_run))
    print('weight sum:', df_run['weight'].sum())
    print('mux min/max:', df_run['mux'].min(), df_run['mux'].max())
    print('muy min/max:', df_run['muy'].min(), df_run['muy'].max())
    print('sigmax min/max:', df_run['sigmax'].min(), df_run['sigmax'].max())
    print('sigmay min/max:', df_run['sigmay'].min(), df_run['sigmay'].max())

    # compute runtime metrics
    metrics_run = plot_sum(df_run, normalize=True, fast=True, return_metrics_only=True)
    print('\nmetrics_run:\n', metrics_run)

    # attempt opt
    df_opt = None
    try:
        sheets = load_all_sheets(str(BASE_WORKBOOK))
        if 'MM configuration' in sheets:
            mm_config = sheets['MM configuration'].copy()
            base_params = _load_base_params_from_workbook(str(BASE_WORKBOOK))
            if mapping is not None:
                base_params['weight'] = base_params['MM #'].astype(int).map(mapping).fillna(0.0)
                # Apply MM_PSF choice to base_params so placement uses the same
                # per-MM PSF parameters as the runtime df.
                try:
                    base_params = apply_mm_psf_choice_to_df(base_params, combo_dict.get('MM_PSF'), aeff_map, rng, standard_presets=standard)
                except Exception:
                    pass
                # Copy minimal runtime centroid/offset fields into base_params
                try:
                    for _col in ('m_rad','m_azi','mux','muy','theta_degrees','theta_position'):
                        if _col in df_run.columns:
                            try:
                                base_params[_col] = base_params['MM #'].astype(int).map(df_run.set_index('MM #')[_col]).fillna(base_params.get(_col, 0.0))
                            except Exception:
                                base_params[_col] = base_params.get(_col, 0.0)
                except Exception:
                    pass
            alignment_by_pos, gravity_by_pos, thermal_by_pos = _load_position_deltas(str(BASE_WORKBOOK))
            # Allow combo overrides: zero out per-position deltas when requested
            def _is_zero_token(v):
                try:
                    if v is None:
                        return False
                    s = str(v).strip()
                    if s == '':
                        return False
                    return float(s) == 0.0
                except Exception:
                    return False

            def _zero_map_for_positions(mm_config_df):
                if 'Position #' in mm_config_df.columns:
                    pos_series = pd.to_numeric(mm_config_df['Position #'], errors='coerce')
                    pos_list = [int(x) for x in pos_series.dropna().astype(int).unique().tolist()]
                else:
                    pos_list = list(range(1, len(mm_config_df) + 1))
                return pos_list

            ck = {k.lower(): k for k in combo_dict.keys()}
            if 'alignment' in ck and _is_zero_token(combo_dict.get(ck['alignment'])):
                pos_list = _zero_map_for_positions(mm_config)
                alignment_by_pos = {p: {'d_align_rad': 0.0, 'd_align_azi': 0.0, 'd_align_z': 0.0, 'd_align_rotz': 0.0} for p in pos_list}
            rawg = None
            for candidate in ('gravity', 'gravity offload', 'grav'):
                if candidate in ck:
                    rawg = combo_dict.get(ck[candidate])
                    break
            if _is_zero_token(rawg):
                pos_list = _zero_map_for_positions(mm_config)
                gravity_by_pos = {p: {'d_grav_x': 0.0, 'd_grav_y': 0.0, 'd_grav_z': 0.0, 'd_grav_rotz': 0.0} for p in pos_list}
            rawt = None
            for candidate in ('thermal', 'therm'):
                if candidate in ck:
                    rawt = combo_dict.get(ck[candidate])
                    break
            if _is_zero_token(rawt):
                pos_list = _zero_map_for_positions(mm_config)
                thermal_by_pos = {p: {'d_therm_x': 0.0, 'd_therm_y': 0.0, 'd_therm_z': 0.0, 'd_therm_rotz': 0.0} for p in pos_list}
            placed_mm = _elliptical_place_mm_config(mm_config, base_params, alignment_by_pos=alignment_by_pos, gravity_by_pos=gravity_by_pos, thermal_by_pos=thermal_by_pos, seed=int(seed))
            df_opt = rebuild_df(base_params, placed_mm)
    except Exception as e:
        print('Error building df_opt:', e)

    if df_opt is None:
        print('\ndf_opt is None')
    else:
        print('\n-- df_opt summary --')
        print('rows:', len(df_opt))
        print('weight sum:', df_opt['weight'].sum())
        print('mux min/max:', df_opt['mux'].min(), df_opt['mux'].max())
        print('muy min/max:', df_opt['muy'].min(), df_opt['muy'].max())
        print('sigmax min/max:', df_opt['sigmax'].min(), df_opt['sigmax'].max())
        print('sigmay min/max:', df_opt['sigmay'].min(), df_opt['sigmay'].max())

    metrics_opt = plot_sum(df_run, normalize=True, fast=True, df_optimized=(df_opt if df_opt is not None else None), return_metrics_only=True)
    print('\nmetrics_opt:\n', metrics_opt)


if __name__ == '__main__':
    inspect_combo(2)
