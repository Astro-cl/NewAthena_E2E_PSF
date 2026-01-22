from pathlib import Path
import pandas as pd
import numpy as np
from tools.run_sensitivity import SENS_PATH, BASE_WORKBOOK, build_aeff_mapping, load_standard_mm_psf_presets, find_aeff_weights_for_choice, load_mm_row_map, apply_mm_psf_choice_to_df
from main import load_gaussians_from_excel, plot_sum
import hashlib

ROOT = Path(__file__).resolve().parent.parent

def main():
    res = pd.read_excel(ROOT / 'Distributions' / 'sensitivity_results.xlsx', engine='openpyxl')
    # pick first row where hew_opt_raw_arcsec is large
    row = res.iloc[0]
    combo = {'A_eff': row['A_eff'], 'MM_PSF': row['MM_PSF']}
    print('Testing combo:', combo)
    aeff_map = build_aeff_mapping()
    aeff_map['mm_row_map'] = load_mm_row_map(BASE_WORKBOOK)
    standard = load_standard_mm_psf_presets(BASE_WORKBOOK)
    h = hashlib.sha1(repr(combo).encode('utf8')).hexdigest()
    seed = int(h[:16],16) % (2**32)
    rng = np.random.default_rng(seed)
    df = load_gaussians_from_excel(str(BASE_WORKBOOK), sheet='MM_PSF')
    mapping = find_aeff_weights_for_choice(combo['A_eff'], aeff_map)
    df['weight'] = df['MM #'].astype(int).map(mapping)
    before_metrics = plot_sum(df, normalize=True, fast=True, return_metrics_only=True)
    print('before hew_best', before_metrics.get('hew_best_arcsec'))
    df2 = apply_mm_psf_choice_to_df(df.copy(), combo['MM_PSF'], aeff_map, rng, standard_presets=standard)
    print('after apply, weight sum', float(df2['weight'].sum()))
    print('sample nonzero MMs:', df2[df2['weight']>0][['MM #','weight']].head(20).to_string(index=False))
    # Now run in-process placement like run_sensitivity
    from optimize_mm_rows import load_all_sheets, _load_base_params_from_workbook, _load_position_deltas, _elliptical_place_mm_config, rebuild_df
    sheets = load_all_sheets(str(BASE_WORKBOOK))
    mm_config = sheets.get('MM configuration')
    base_params = _load_base_params_from_workbook(str(BASE_WORKBOOK))
    # override weights
    mapped = base_params['MM #'].astype(int).map(mapping).fillna(0.0)
    base_params['weight'] = mapped
    # Apply MM_PSF choice to base_params so placement uses the same
    # per-MM PSF parameters as the runtime-applied df2.
    try:
        base_params = apply_mm_psf_choice_to_df(base_params, combo.get('MM_PSF'), aeff_map, rng, standard_presets=standard)
    except Exception:
        pass
    # Copy minimal runtime centroid/offset fields into base_params so rebuild_df preserves centers
    try:
        for _col in ('m_rad','m_azi','mux','muy','theta_degrees','theta_position'):
            if _col in df2.columns:
                try:
                    base_params[_col] = base_params['MM #'].astype(int).map(df2.set_index('MM #')[_col]).fillna(base_params.get(_col, 0.0))
                except Exception:
                    base_params[_col] = base_params.get(_col, 0.0)
    except Exception:
        pass
    alignment_by_pos, gravity_by_pos, thermal_by_pos = _load_position_deltas(str(BASE_WORKBOOK))
    # Allow Sensitivity combo to request zeroing of position deltas
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

    ck = {k.lower(): k for k in combo.keys()}
    if 'alignment' in ck and _is_zero_token(combo.get(ck['alignment'])):
        pos_list = _zero_map_for_positions(mm_config)
        alignment_by_pos = {p: {'d_align_rad': 0.0, 'd_align_azi': 0.0, 'd_align_z': 0.0, 'd_align_rotz': 0.0} for p in pos_list}
    rawg = None
    for candidate in ('gravity', 'gravity offload', 'grav'):
        if candidate in ck:
            rawg = combo.get(ck[candidate])
            break
    if _is_zero_token(rawg):
        pos_list = _zero_map_for_positions(mm_config)
        gravity_by_pos = {p: {'d_grav_x': 0.0, 'd_grav_y': 0.0, 'd_grav_z': 0.0, 'd_grav_rotz': 0.0} for p in pos_list}
    rawt = None
    for candidate in ('thermal', 'therm'):
        if candidate in ck:
            rawt = combo.get(ck[candidate])
            break
    if _is_zero_token(rawt):
        pos_list = _zero_map_for_positions(mm_config)
        thermal_by_pos = {p: {'d_therm_x': 0.0, 'd_therm_y': 0.0, 'd_therm_z': 0.0, 'd_therm_rotz': 0.0} for p in pos_list}
    seed_i = int(seed)
    placed_mm = _elliptical_place_mm_config(mm_config, base_params, alignment_by_pos=alignment_by_pos, gravity_by_pos=gravity_by_pos, thermal_by_pos=thermal_by_pos, seed=seed_i)
    final_df = rebuild_df(base_params, placed_mm)
    print('df_opt weight sum', float(final_df['weight'].sum()))
    post_metrics = plot_sum(df2, normalize=True, fast=True, df_optimized=final_df, return_metrics_only=True)
    print('post metrics (with df_optimized):', post_metrics)
    post_metrics_raw = plot_sum(df2, normalize=True, fast=True, df_optimized=final_df, return_metrics_only=True)
    print('post raw hew_opt_arcsec:', post_metrics_raw.get('hew_opt_arcsec'))
    # inspect df_opt columns sample
    print('df_opt sample rows:', final_df[['MM #','mux','muy','sigmax','sigmay']].head(20).to_string(index=False))

if __name__=='__main__':
    main()
