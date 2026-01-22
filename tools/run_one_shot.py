import sys, pathlib
ROOT=pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0,str(ROOT))
from tools.run_sensitivity import SENS_PATH, BASE_WORKBOOK, OUT_XLSX, build_aeff_mapping, load_standard_mm_psf_presets, load_mm_row_map, find_aeff_weights_for_choice, apply_mm_psf_choice_to_df
from main import load_gaussians_from_excel, plot_sum
import pandas as pd, itertools, hashlib, numpy as np, re

sens = pd.read_excel(SENS_PATH, sheet_name=0, engine='openpyxl', header=0)
param_options = {}
for col in sens.columns:
    vals = sens[col].dropna().astype(str).map(str.strip)
    vals = [v for v in vals if v not in ['', '-', 'NaN', 'nan', 'None']]
    if vals:
        param_options[col] = vals

# expand A_eff [row#]
aeff_map = build_aeff_mapping()
aeff_map['mm_row_map'] = load_mm_row_map(BASE_WORKBOOK)
if 'A_eff' in param_options:
    vals = param_options['A_eff']
    expanded = []
    row_token_re = re.compile(r'^(.*?)\s*\[row#\]\s*$', re.IGNORECASE)
    if aeff_map.get('mm_row_map'):
        row_keys = sorted(aeff_map['mm_row_map'].keys())
    else:
        row_keys = []
    for v in vals:
        m = row_token_re.match(v)
        if m and row_keys:
            base = m.group(1).strip()
            for r in row_keys:
                expanded.append(f"{base} [row{r}]")
        else:
            expanded.append(v)
    param_options['A_eff'] = expanded

keys = list(param_options.keys())
combos = list(itertools.product(*(param_options[k] for k in keys)))
standard = load_standard_mm_psf_presets(BASE_WORKBOOK)
results = []
for idx, combo in enumerate(combos, start=1):
    combo_dict = dict(zip(keys, combo))
    h = hashlib.sha1(repr(combo_dict).encode('utf8')).hexdigest()
    seed = int(h[:16], 16) % (2**32)
    rng = np.random.default_rng(seed)
    df = load_gaussians_from_excel(str(BASE_WORKBOOK), sheet='MM_PSF')
    if 'A_eff' in combo_dict:
        try:
            mapping = find_aeff_weights_for_choice(combo_dict['A_eff'], aeff_map)
            df['weight'] = df['MM #'].astype(int).map(mapping)
        except Exception:
            pass
    if 'MM_PSF' in combo_dict:
        df = apply_mm_psf_choice_to_df(df, combo_dict['MM_PSF'], aeff_map, rng, standard_presets=standard)
    # capture total weight and distribution diversity for diagnostics
    try:
        total_weight = float(df['weight'].sum())
    except Exception:
        total_weight = None
    try:
        n_unique_dists = int(df.get('distribution', pd.Series([])).astype(str).str.lower().nunique())
    except Exception:
        n_unique_dists = None

    metrics = plot_sum(df, normalize=True, fast=True, return_metrics_only=True)
    keep = ['hew_origin_arcsec','hew_best_arcsec','hew_opt_arcsec','eef90_origin_arcsec','eef90_best_arcsec','eef90_opt_arcsec']
    trimmed = {k: (metrics.get(k) if isinstance(metrics, dict) else None) for k in keep}
    # Attempt in-process placement to compute raw post-placement optimized metrics
    post_metrics_raw = None
    df_opt = None
    try:
        from optimize_mm_rows import load_all_sheets, _load_base_params_from_workbook, _load_position_deltas, _elliptical_place_mm_config, rebuild_df
        sheets = load_all_sheets(str(BASE_WORKBOOK))
        if 'MM configuration' in sheets:
            mm_config = sheets['MM configuration'].copy()
            base_params = _load_base_params_from_workbook(str(BASE_WORKBOOK))
            try:
                    if 'mapping' in locals() and mapping is not None:
                        mapped = base_params['MM #'].astype(int).map(mapping).fillna(0.0)
                        base_params['weight'] = mapped
                        # Do NOT copy runtime per-MM fields into base_params; rebuild_df
                        # will preserve runtime values where available.
            except Exception:
                pass
            alignment_by_pos, gravity_by_pos, thermal_by_pos = _load_position_deltas(str(BASE_WORKBOOK))
            # Respect sensitivity combo overrides: if Alignment/Gravity/Thermal set to zero, zero those deltas.
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

            combo_keys = {k.lower(): k for k in combo_dict.keys()}
            if 'alignment' in combo_keys or 'align' in combo_keys:
                raw = combo_dict.get(combo_keys.get('alignment', combo_keys.get('align')))
                if _is_zero_token(raw):
                    pos_list = _zero_map_for_positions(mm_config)
                    alignment_by_pos = {p: {'d_align_rad': 0.0, 'd_align_azi': 0.0, 'd_align_z': 0.0, 'd_align_rotz': 0.0} for p in pos_list}
            if 'gravity' in combo_keys or 'gravity offload' in combo_keys or 'grav' in combo_keys:
                raw = None
                for candidate in ('gravity', 'gravity offload', 'grav'):
                    if candidate in combo_keys:
                        raw = combo_dict.get(combo_keys[candidate])
                        break
                if _is_zero_token(raw):
                    pos_list = _zero_map_for_positions(mm_config)
                    gravity_by_pos = {p: {'d_grav_x': 0.0, 'd_grav_y': 0.0, 'd_grav_z': 0.0, 'd_grav_rotz': 0.0} for p in pos_list}
            if 'thermal' in combo_keys or 'therm' in combo_keys:
                raw = None
                for candidate in ('thermal', 'therm'):
                    if candidate in combo_keys:
                        raw = combo_dict.get(combo_keys[candidate])
                        break
                if _is_zero_token(raw):
                    pos_list = _zero_map_for_positions(mm_config)
                    thermal_by_pos = {p: {'d_therm_x': 0.0, 'd_therm_y': 0.0, 'd_therm_z': 0.0, 'd_therm_rotz': 0.0} for p in pos_list}
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
            try:
                post_metrics_raw = plot_sum(df, normalize=True, fast=True, df_optimized=df_opt, return_metrics_only=True)
            except Exception:
                post_metrics_raw = None
    except Exception:
        post_metrics_raw = None
    combo_out = dict(combo_dict)
    # Parse MM_PSF details: symmetry and distribution type
    mmpsf_raw = str(combo_out.get('MM_PSF', '') or '')
    def _parse_sym(s):
        if re.search(r'\bAsym\b|Asymmetric', s, re.IGNORECASE):
            return 'Asym'
        if re.search(r'\bSym\b|Symmetric', s, re.IGNORECASE):
            return 'Sym'
        return ''
    def _mmpsf_type(s):
        sl = s.lower()
        if 'voigt' in sl or 'pseudo-voigt' in sl:
            return 'Pseudo-Voigt'
        if 'gaussian' in sl:
            return 'Gaussian'
        # treat named presets or other tokens as Custom
        if standard and any(k.lower() in sl for k in standard.keys()):
            return 'Custom'
        return 'Custom'
    combo_out['MM_PSF_symmetry'] = _parse_sym(mmpsf_raw)
    # Default to 'Asym' when a Voigt/pseudo-Voigt token is present and no explicit symmetry
    if ('voigt' in mmpsf_raw.lower() or 'pseudo-voigt' in mmpsf_raw.lower()) and not combo_out['MM_PSF_symmetry']:
        combo_out['MM_PSF_symmetry'] = 'Asym'
    combo_out['MM_PSF_type'] = _mmpsf_type(mmpsf_raw)
    combo_out['total_weight'] = total_weight
    combo_out['n_unique_distributions'] = n_unique_dists
    # attach raw optimized metrics and placement flag when available
    try:
        combo_out['hew_opt_raw_arcsec'] = post_metrics_raw.get('hew_opt_arcsec') if isinstance(post_metrics_raw, dict) else None
        combo_out['eef90_opt_raw_arcsec'] = post_metrics_raw.get('eef90_opt_arcsec') if isinstance(post_metrics_raw, dict) else None
        combo_out['placement_improved'] = (combo_out['hew_opt_raw_arcsec'] is not None and combo_out.get('hew_best_arcsec') is not None and combo_out['hew_opt_raw_arcsec'] < combo_out.get('hew_best_arcsec'))
    except Exception:
        combo_out['hew_opt_raw_arcsec'] = None
        combo_out['eef90_opt_raw_arcsec'] = None
        combo_out['placement_improved'] = False
    row_num = '-'
    aeff_val = combo_out.get('A_eff')
    if aeff_val is not None:
        m = re.match(r'^(.*?)\s*\[row(\d+)\]\s*$', str(aeff_val), re.IGNORECASE)
        if m:
            combo_out['A_eff'] = m.group(1).strip()
            try:
                row_num = int(m.group(2))
            except Exception:
                row_num = '-'
    combo_out['row'] = row_num
    results.append({**combo_out, **trimmed})

# save and normalize A_eff/row at DataFrame level
df_res = pd.DataFrame(results)
if 'A_eff' in df_res.columns:
    def _extract_row(x):
        try:
            if pd.isna(x):
                return '-'
            m = re.match(r'^(.*?)\s*\[row(\d+)\]\s*$', str(x), re.IGNORECASE)
            if m:
                return int(m.group(2))
        except Exception:
            return '-'
        return '-'
    df_res['row'] = df_res['A_eff'].map(_extract_row)
    def _strip_row(x):
        try:
            if pd.isna(x):
                return x
            m = re.match(r'^(.*?)\s*\[row(\d+)\]\s*$', str(x), re.IGNORECASE)
            if m:
                return m.group(1).strip()
        except Exception:
            return x
        return x
    df_res['A_eff'] = df_res['A_eff'].map(_strip_row)
if 'row' not in df_res.columns:
    df_res['row'] = '-'
# Reorder so 'row' is column B if possible
cols = list(df_res.columns)
# Ensure MM_PSF_symmetry and MM_PSF_type are placed after MM_PSF (so Excel col D shows symmetry)
if 'MM_PSF' in cols:
    # build desired order: A_eff, row, MM_PSF, MM_PSF_symmetry, MM_PSF_type, then rest
    pre = []
    if 'A_eff' in cols:
        pre.append('A_eff')
    if 'row' in cols:
        pre.append('row')
    pre.append('MM_PSF')
    if 'MM_PSF_symmetry' in cols:
        pre.append('MM_PSF_symmetry')
    if 'MM_PSF_type' in cols:
        pre.append('MM_PSF_type')
    remaining = [c for c in cols if c not in pre]
    cols = pre + remaining
if 'A_eff' in cols and 'row' in cols:
        cols = [c for c in cols if c not in ('A_eff','row')]
        cols = ['A_eff','row'] + cols
        df_res = df_res[cols]

df_res.to_excel(OUT_XLSX, index=False)
print('Wrote', OUT_XLSX)
print('Columns:', df_res.columns.tolist())
print(df_res[['A_eff','row']].head(20).to_string(index=False))
