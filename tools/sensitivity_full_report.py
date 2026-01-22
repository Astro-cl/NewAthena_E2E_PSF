"""
Produce a per-combo metrics comparison report for the sensitivity sweep.
For each combo in Distributions/Sensitivity.xlsx:
 - apply A_eff mapping and MM_PSF choice to base MM_PSF
 - run in-process placement to get df_opt (if placement available)
 - compute metrics via plot_sum for runtime df and df_opt
 - write per-combo row to Figures/sensitivity_full_report_<ts>.csv
Also writes detailed dumps for combos that reference pseudo-voigt.
"""
from pathlib import Path
import sys, time, hashlib, itertools, json
import pandas as pd
import numpy as np

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from tools.run_sensitivity import SENS_PATH, BASE_WORKBOOK, build_aeff_mapping, load_standard_mm_psf_presets, load_mm_row_map, find_aeff_weights_for_choice, apply_mm_psf_choice_to_df
from main import load_gaussians_from_excel, plot_sum
from optimize_mm_rows import load_all_sheets, _load_base_params_from_workbook, _load_position_deltas, _elliptical_place_mm_config, rebuild_df

OUTDIR = ROOT / 'Figures'
OUTDIR.mkdir(parents=True, exist_ok=True)


def collect_param_options(sens_path):
    sens = pd.read_excel(sens_path, sheet_name=0, engine='openpyxl', header=0)
    param_options = {}
    for col in sens.columns:
        vals = sens[col].dropna().astype(str).map(str.strip)
        vals = [v for v in vals if v not in ['', '-', 'NaN', 'nan', 'None']]
        if vals:
            param_options[col] = vals
    return param_options


def run():
    param_options = collect_param_options(SENS_PATH)
    if not param_options:
        print('No sensitivity options found; aborting')
        return
    keys = list(param_options.keys())
    combos = list(itertools.product(*(param_options[k] for k in keys)))
    print(f'Found {len(combos)} combos')

    aeff_map = build_aeff_mapping()
    aeff_map['mm_row_map'] = load_mm_row_map(BASE_WORKBOOK)
    standard = load_standard_mm_psf_presets(BASE_WORKBOOK)

    rows = []
    ts = int(time.time())
    for idx, combo in enumerate(combos, start=1):
        combo_dict = dict(zip(keys, combo))
        h = hashlib.sha1(repr(combo_dict).encode('utf8')).hexdigest()
        seed = int(h[:16], 16) % (2**32)
        rng = np.random.default_rng(seed)

        df_run = load_gaussians_from_excel(str(BASE_WORKBOOK), sheet='MM_PSF')
        mapping = None
        if 'A_eff' in combo_dict:
            try:
                mapping = find_aeff_weights_for_choice(combo_dict['A_eff'], aeff_map)
                df_run['weight'] = df_run['MM #'].astype(int).map(mapping)
            except Exception as e:
                print('  Warning: A_eff mapping failed:', e)
        if 'MM_PSF' in combo_dict:
            try:
                df_run = apply_mm_psf_choice_to_df(df_run, combo_dict['MM_PSF'], aeff_map, rng, standard_presets=standard)
            except Exception as e:
                print('  Warning: apply_mm_psf failed:', e)

        # compute runtime metrics
        try:
            metrics_run = plot_sum(df_run, normalize=True, fast=True, return_metrics_only=True)
        except Exception as e:
            metrics_run = {'error_plot_sum': str(e)}

        # attempt in-process placement to derive optimized df
        df_opt = None
        try:
            sheets = load_all_sheets(str(BASE_WORKBOOK))
            if 'MM configuration' in sheets:
                mm_config = sheets['MM configuration'].copy()
                base_params = _load_base_params_from_workbook(str(BASE_WORKBOOK))
                if mapping is not None:
                    base_params['weight'] = base_params['MM #'].astype(int).map(mapping).fillna(0.0)
                    # Apply MM_PSF choice to base_params so in-process placement uses
                    # the same per-MM PSF parameters as the runtime `df_run`.
                    try:
                        base_params = apply_mm_psf_choice_to_df(base_params, combo_dict.get('MM_PSF'), aeff_map, rng, standard_presets=standard)
                    except Exception:
                        pass
                    # Copy minimal runtime centroid/offset fields into base_params so rebuild_df
                    # can preserve runtime centers during placement.
                    try:
                        for _col in ('m_rad','m_azi','mux','muy','theta_degrees','theta_position'):
                            if _col in df_run.columns:
                                try:
                                    base_params[_col] = base_params['MM #'].astype(int).map(df_run.set_index('MM #')[_col]).fillna(base_params.get(_col,0.0))
                                except Exception:
                                    base_params[_col] = base_params.get(_col,0.0)
                    except Exception:
                        pass
                alignment_by_pos, gravity_by_pos, thermal_by_pos = _load_position_deltas(str(BASE_WORKBOOK))
                # Allow Sensitivity combo to zero-out position deltas (Alignment/Gravity/Thermal)
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
                if 'alignment' in ck:
                    raw = combo_dict.get(ck['alignment'])
                    if _is_zero_token(raw):
                        pos_list = _zero_map_for_positions(mm_config)
                        alignment_by_pos = {p: {'d_align_rad': 0.0, 'd_align_azi': 0.0, 'd_align_z': 0.0, 'd_align_rotz': 0.0} for p in pos_list}
                # gravity
                rawg = None
                for candidate in ('gravity', 'gravity offload', 'grav'):
                    if candidate in ck:
                        rawg = combo_dict.get(ck[candidate])
                        break
                if _is_zero_token(rawg):
                    pos_list = _zero_map_for_positions(mm_config)
                    gravity_by_pos = {p: {'d_grav_x': 0.0, 'd_grav_y': 0.0, 'd_grav_z': 0.0, 'd_grav_rotz': 0.0} for p in pos_list}
                # thermal
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
            df_opt = None

        try:
            metrics_opt = plot_sum(df_run, normalize=True, fast=True, df_optimized=(df_opt if df_opt is not None else None), return_metrics_only=True)
        except Exception as e:
            metrics_opt = {'error_plot_sum_opt': str(e)}

        mismatches = {}
        # compare key metrics
        keys_keep = ['hew_origin_arcsec','hew_best_arcsec','hew_opt_arcsec','eef90_origin_arcsec','eef90_best_arcsec','eef90_opt_arcsec']
        for k in keys_keep:
            v_run = metrics_run.get(k) if isinstance(metrics_run, dict) else None
            v_opt = metrics_opt.get(k) if isinstance(metrics_opt, dict) else None
            if v_run is None and v_opt is None:
                delta = None
            else:
                try:
                    delta = (v_opt - v_run) if (v_run is not None and v_opt is not None) else None
                except Exception:
                    delta = None
            mismatches[f'delta_{k}'] = delta

        row = {
            'idx': idx,
            'combo': json.dumps(combo_dict),
            **{f'run_{k}': metrics_run.get(k) if isinstance(metrics_run, dict) else None for k in keys_keep},
            **{f'opt_{k}': metrics_opt.get(k) if isinstance(metrics_opt, dict) else None for k in keys_keep},
            **mismatches,
        }
        rows.append(row)

        # detailed dump for pseudo-voigt combos
        mmpsf = str(combo_dict.get('MM_PSF',''))
        if 'voigt' in mmpsf.lower() or 'pseudo-voigt' in mmpsf.lower() or any(str(v).lower().strip() in str(mmpsf).lower() for v in (list(standard.keys()) if standard else [])):
            fn = OUTDIR / f'combo_{idx}_details_{ts}.csv'
            try:
                # write runtime and opt first 50 rows
                dump = []
                df_run_small = df_run.copy()
                df_run_small['source'] = 'runtime'
                dump.append(df_run_small.head(200))
                if df_opt is not None:
                    df_opt_small = df_opt.copy()
                    df_opt_small['source'] = 'opt'
                    dump.append(df_opt_small.head(200))
                pd.concat(dump, ignore_index=True).to_csv(fn, index=False)
            except Exception as e:
                print('  Warning: could not write details for combo', idx, e)

    outp = OUTDIR / f'sensitivity_full_report_{ts}.csv'
    pd.DataFrame(rows).to_csv(outp, index=False)
    print('Wrote', outp)


if __name__ == '__main__':
    run()
