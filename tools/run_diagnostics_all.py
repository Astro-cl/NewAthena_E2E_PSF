from pathlib import Path
import sys, json
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import pandas as pd
import numpy as np
import hashlib

from tools.run_sensitivity import SENS_PATH, BASE_WORKBOOK, build_aeff_mapping, load_standard_mm_psf_presets, find_aeff_weights_for_choice, load_mm_row_map, apply_mm_psf_choice_to_df
from main import load_gaussians_from_excel, plot_sum

from optimize_mm_rows import load_all_sheets, _load_base_params_from_workbook, _load_position_deltas, _elliptical_place_mm_config, rebuild_df


OUT_DIR = ROOT / 'Figures' / 'diagnostics'
OUT_DIR.mkdir(parents=True, exist_ok=True)


def run_one(idx, row):
    combo = {'A_eff': row.get('A_eff'), 'MM_PSF': row.get('MM_PSF')}
    print(f'[{idx}] combo:', combo)
    aeff_map = build_aeff_mapping()
    aeff_map['mm_row_map'] = load_mm_row_map(BASE_WORKBOOK)
    standard = load_standard_mm_psf_presets(BASE_WORKBOOK)

    h = hashlib.sha1(repr(combo).encode('utf8')).hexdigest()
    seed = int(h[:16], 16) % (2**32)
    rng = np.random.default_rng(seed)

    df_run = load_gaussians_from_excel(str(BASE_WORKBOOK), sheet='MM_PSF')
    try:
        mapping = find_aeff_weights_for_choice(combo['A_eff'], aeff_map)
        df_run['weight'] = df_run['MM #'].astype(int).map(mapping)
    except Exception as e:
        print('  A_eff mapping failed:', e)

    df_run = apply_mm_psf_choice_to_df(df_run, combo.get('MM_PSF'), aeff_map, rng, standard_presets=standard)

    # save df_run
    outsub = OUT_DIR / f'combo_{idx}'
    outsub.mkdir(parents=True, exist_ok=True)
    df_run.to_csv(outsub / 'df_run.csv', index=False)

    # compute before metrics
    try:
        metrics_pre = plot_sum(df_run, normalize=True, fast=True, return_metrics_only=True)
    except Exception as e:
        metrics_pre = {'error': str(e)}

    # attempt in-process placement
    df_opt = None
    post_metrics_raw = None
    try:
        sheets = load_all_sheets(str(BASE_WORKBOOK))
        mm_config = sheets.get('MM configuration')
        base_params = _load_base_params_from_workbook(str(BASE_WORKBOOK))
        alignment_by_pos, gravity_by_pos, thermal_by_pos = _load_position_deltas(str(BASE_WORKBOOK))
        # Allow Sensitivity combo to disable position deltas (zero them)
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

        combo_keys = {k.lower(): k for k in combo.keys()}
        if 'alignment' in combo_keys or 'align' in combo_keys:
            raw = combo.get(combo_keys.get('alignment', combo_keys.get('align')))
            if _is_zero_token(raw):
                pos_list = _zero_map_for_positions(mm_config)
                alignment_by_pos = {p: {'d_align_rad': 0.0, 'd_align_azi': 0.0, 'd_align_z': 0.0, 'd_align_rotz': 0.0} for p in pos_list}
        # gravity
        rawg = None
        for candidate in ('gravity', 'gravity offload', 'grav'):
            if candidate in combo_keys:
                rawg = combo.get(combo_keys[candidate])
                break
        if _is_zero_token(rawg):
            pos_list = _zero_map_for_positions(mm_config)
            gravity_by_pos = {p: {'d_grav_x': 0.0, 'd_grav_y': 0.0, 'd_grav_z': 0.0, 'd_grav_rotz': 0.0} for p in pos_list}
        # thermal
        rawt = None
        for candidate in ('thermal', 'therm'):
            if candidate in combo_keys:
                rawt = combo.get(combo_keys[candidate])
                break
        if _is_zero_token(rawt):
            pos_list = _zero_map_for_positions(mm_config)
            thermal_by_pos = {p: {'d_therm_x': 0.0, 'd_therm_y': 0.0, 'd_therm_z': 0.0, 'd_therm_rotz': 0.0} for p in pos_list}
        # override weights
        try:
            mapped = base_params['MM #'].astype(int).map(mapping).fillna(0.0)
            base_params['weight'] = mapped
        except Exception:
            pass
        # Apply the MM_PSF choice to base_params so placement uses the same
        # per-MM PSF parameters as the runtime-sampled `df_run`.
        try:
            base_params = apply_mm_psf_choice_to_df(base_params, combo.get('MM_PSF'), aeff_map, rng, standard_presets=standard)
        except Exception:
            # best-effort only; placement will proceed with whatever base_params contains
            pass
        # Ensure runtime centroids/offsets (mux/muy/m_rad/m_azi) are present in base_params
        # so rebuild_df can preserve them. Only copy minimal center fields (not sigmas).
        try:
            for _col in ('m_rad', 'm_azi', 'mux', 'muy', 'theta_degrees', 'theta_position'):
                if _col in df_run.columns:
                    try:
                        base_params[_col] = base_params['MM #'].astype(int).map(df_run.set_index('MM #')[_col]).fillna(base_params.get(_col, 0.0))
                    except Exception:
                        base_params[_col] = base_params.get(_col, 0.0)
        except Exception:
            pass

        placed_mm = _elliptical_place_mm_config(
            mm_config,
            base_params,
            alignment_by_pos=alignment_by_pos,
            gravity_by_pos=gravity_by_pos,
            thermal_by_pos=thermal_by_pos,
            seed=int(seed),
        )
        final_df = rebuild_df(base_params, placed_mm)
        df_opt = final_df
        final_df.to_csv(outsub / 'df_opt.csv', index=False)
        try:
            post_metrics_raw = plot_sum(df_run, normalize=True, fast=True, df_optimized=df_opt, return_metrics_only=True)
        except Exception as e:
            post_metrics_raw = {'error': str(e)}
    except Exception as e:
        print('  in-process placement failed:', e)

    # Write metrics
    summary = {'pre': metrics_pre, 'post_raw': post_metrics_raw}
    with open(outsub / 'metrics.json', 'w') as fh:
        json.dump(summary, fh, default=str, indent=2)

    return idx, combo, metrics_pre, post_metrics_raw


def main():
    res_path = ROOT / 'Distributions' / 'sensitivity_results.xlsx'
    if not res_path.exists():
        print('sensitivity_results.xlsx not found')
        return
    res = pd.read_excel(res_path, engine='openpyxl')
    # select troubling combos: placement_improved False or hew_opt_raw_arcsec > hew_best_arcsec
    to_run = []
    for i, row in res.iterrows():
        idx = i + 1
        improved = row.get('placement_improved')
        raw_opt = row.get('hew_opt_raw_arcsec')
        best = row.get('hew_best_arcsec')
        if improved is False or (pd.notna(raw_opt) and pd.notna(best) and raw_opt > best * 1.0):
            to_run.append((idx, row))

    print(f'Found {len(to_run)} combos to diagnose')
    results = []
    for idx, row in to_run:
        try:
            out = run_one(idx, row)
            results.append(out)
        except Exception as e:
            print('Failed combo', idx, e)

    # write summary CSV
    rows = []
    for idx, combo, pre, post in results:
        rows.append({
            'idx': idx,
            'A_eff': combo.get('A_eff'),
            'MM_PSF': combo.get('MM_PSF'),
            'hew_best_run': (pre.get('hew_best_arcsec') if isinstance(pre, dict) else None),
            'hew_opt_raw': (post.get('hew_opt_arcsec') if isinstance(post, dict) else None),
        })
    pd.DataFrame(rows).to_csv(OUT_DIR / 'diagnostics_summary.csv', index=False)
    print('Wrote diagnostics to', OUT_DIR)


if __name__ == '__main__':
    main()
