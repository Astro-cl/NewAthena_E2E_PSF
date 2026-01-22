"""
Validate mux/muy preservation across all sensitivity combinations.
Writes a per-combo report to `Figures/validate_sensitivity_mux_muy_<ts>.csv`.
Run: python3 tools/validate_sensitivity_mux_muy.py
"""
from pathlib import Path
import sys
import time
import itertools
import hashlib
import pandas as pd
import numpy as np
import json

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from tools.run_sensitivity import SENS_PATH, BASE_WORKBOOK, build_aeff_mapping, load_standard_mm_psf_presets, load_mm_row_map, find_aeff_weights_for_choice, apply_mm_psf_choice_to_df
from main import load_gaussians_from_excel
from optimize_mm_rows import load_all_sheets, _load_base_params_from_workbook, rebuild_df

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


def run_validation():
    param_options = collect_param_options(SENS_PATH)
    if not param_options:
        print('No sensitivity options found; aborting')
        return
    keys = list(param_options.keys())
    combos = list(itertools.product(*(param_options[k] for k in keys)))
    print(f'Validating {len(combos)} combos')

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

        try:
            df_runtime = load_gaussians_from_excel(str(BASE_WORKBOOK), sheet='MM_PSF')
        except Exception as e:
            rows.append({'idx': idx, 'combo': combo_dict, 'error': f'load_base_failed: {e}'})
            continue

        mapping = None
        if 'A_eff' in combo_dict:
            try:
                mapping = find_aeff_weights_for_choice(combo_dict['A_eff'], aeff_map)
                df_runtime['weight'] = df_runtime['MM #'].astype(int).map(mapping)
            except Exception as e:
                rows.append({'idx': idx, 'combo': combo_dict, 'error': f'aeff_map_failed: {e}'})
                continue

        if 'MM_PSF' in combo_dict:
            try:
                df_runtime = apply_mm_psf_choice_to_df(df_runtime, combo_dict['MM_PSF'], aeff_map, rng, standard_presets=standard)
            except Exception as e:
                rows.append({'idx': idx, 'combo': combo_dict, 'error': f'apply_mm_psf_failed: {e}'})
                continue

        # Build base_params and mm_config like run_sensitivity
        try:
            sheets = load_all_sheets(str(BASE_WORKBOOK))
            if 'MM configuration' not in sheets:
                rows.append({'idx': idx, 'combo': combo_dict, 'error': 'mm_config_missing'})
                continue
            mm_config = sheets['MM configuration'].copy()
            base_params = _load_base_params_from_workbook(str(BASE_WORKBOOK))
            if mapping is not None:
                mapped = base_params['MM #'].astype(int).map(mapping).fillna(0.0)
                base_params['weight'] = mapped
                # Apply MM_PSF choice to base_params so placement uses the same
                # per-MM PSF parameters as the runtime df.
                try:
                    base_params = apply_mm_psf_choice_to_df(base_params, combo_dict.get('MM_PSF'), aeff_map, rng, standard_presets=standard)
                except Exception:
                    pass
                # Copy minimal runtime centroid/offset fields into base_params
                try:
                    for _col in ('m_rad', 'm_azi', 'mux', 'muy', 'theta_degrees', 'theta_position'):
                        if _col in df_runtime.columns:
                            try:
                                base_params[_col] = base_params['MM #'].astype(int).map(df_runtime.set_index('MM #')[_col]).fillna(base_params.get(_col, 0.0))
                            except Exception:
                                base_params[_col] = base_params.get(_col, 0.0)
                except Exception:
                    pass
        except Exception as e:
            rows.append({'idx': idx, 'combo': combo_dict, 'error': f'build_base_failed: {e}'})
            continue

        try:
            final_df = rebuild_df(base_params, mm_config)
        except Exception as e:
            rows.append({'idx': idx, 'combo': combo_dict, 'error': f'rebuild_failed: {e}'})
            continue

        # Compare per-MM values: for MM present in df_runtime
        tol = 1e-15
        mismatches = 0
        samples = []
        runtime_map = {}
        try:
            runtime_map = df_runtime.set_index(df_runtime['MM #'].astype(int))[['mux', 'muy']].to_dict(orient='index')
        except Exception:
            # fallback align by index
            runtime_map = None
        for _, rowp in final_df.iterrows():
            mm = int(rowp['MM #'])
            if runtime_map is None or mm not in runtime_map:
                continue
            rvals = runtime_map[mm]
            mux_r = float(rvals.get('mux', float('nan')))
            muy_r = float(rvals.get('muy', float('nan')))
            mux_f = float(rowp.get('mux', float('nan')))
            muy_f = float(rowp.get('muy', float('nan')))
            def changed(a, b):
                if (pd.isna(a) and not pd.isna(b)) or (pd.isna(b) and not pd.isna(a)):
                    return True
                if pd.isna(a) and pd.isna(b):
                    return False
                try:
                    return abs(float(a) - float(b)) > tol
                except Exception:
                    return True
            if changed(mux_r, mux_f) or changed(muy_r, muy_f):
                mismatches += 1
                if len(samples) < 5:
                    samples.append({'MM': mm, 'mux_runtime': mux_r, 'muy_runtime': muy_r, 'mux_final': mux_f, 'muy_final': muy_f})

        rows.append({'idx': idx, 'combo': combo_dict, 'mismatches': mismatches, 'samples': samples})

    outf = OUTDIR / f'validate_sensitivity_mux_muy_{ts}.csv'
    # Flatten rows to CSV-friendly
    flat = []
    for r in rows:
        entry = {k: v for k, v in r.items() if k not in ('combo', 'samples')}
        entry['combo'] = json.dumps(r.get('combo'))
        entry['samples'] = json.dumps(r.get('samples'))
        flat.append(entry)
    pd.DataFrame(flat).to_csv(outf, index=False)
    print('Wrote', outf)


if __name__ == '__main__':
    run_validation()
