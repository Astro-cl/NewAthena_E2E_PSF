from pathlib import Path
import sys, itertools, hashlib
import pandas as pd
import numpy as np

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from tools.run_sensitivity import SENS_PATH, BASE_WORKBOOK, build_aeff_mapping, load_standard_mm_psf_presets, find_aeff_weights_for_choice, load_mm_row_map
from main import load_gaussians_from_excel, plot_sum
from tools.run_sensitivity import OUT_XLSX


def main():
    print('Loading sensitivity workbook and mapping...')
    sens = pd.read_excel(SENS_PATH, sheet_name=0, engine='openpyxl', header=0)
    param_options = {}
    for col in sens.columns:
        vals = sens[col].dropna().astype(str).map(str.strip)
        vals = [v for v in vals if v not in ['', '-', 'NaN', 'nan', 'None']]
        if vals:
            param_options[col] = vals

    # expand A_eff [row#] tokens like run_one_shot
    import re
    aeff_map = build_aeff_mapping()
    # populate mm_row_map like run_one_shot does
    try:
        aeff_map['mm_row_map'] = load_mm_row_map(BASE_WORKBOOK)
    except Exception:
        pass
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

    # reuse aeff_map (already populated above with mm_row_map)
    standard = load_standard_mm_psf_presets(BASE_WORKBOOK)

    target = 'pseudo-voigt'
    findings = []
    for idx, combo in enumerate(combos, start=1):
        combo_dict = dict(zip(keys, combo))
        mmpsf = str(combo_dict.get('MM_PSF',''))
        if target in mmpsf.lower():
            aeff = combo_dict.get('A_eff','')
            # extract row number if present
            import re
            m = re.match(r'^(.*?)\s*\[row(\d+)\]\s*$', str(aeff), re.IGNORECASE)
            rownum = int(m.group(2)) if m else None
            if rownum in (6,7):
                findings.append((idx, combo_dict, rownum))

    if not findings:
        print('No matching combos for rows 6 or 7 found.')
        return

    for idx, combo_dict, rownum in findings:
        print('\n=== Combo idx', idx, 'row', rownum, '===')
        print(combo_dict)
        # build deterministic rng
        h = hashlib.sha1(repr(combo_dict).encode('utf8')).hexdigest()
        seed = int(h[:16], 16) % (2**32)
        rng = np.random.default_rng(seed)

        df = load_gaussians_from_excel(str(BASE_WORKBOOK), sheet='MM_PSF')
        mapping = None
        try:
            mapping = find_aeff_weights_for_choice(combo_dict.get('A_eff',''), aeff_map)
            df['weight'] = df['MM #'].astype(int).map(mapping)
        except Exception as e:
            print('  A_eff mapping failed:', e)
        df = df.copy()
        df = __import__('tools.run_sensitivity', fromlist=['apply_mm_psf_choice_to_df']).apply_mm_psf_choice_to_df(df, combo_dict.get('MM_PSF'), aeff_map, rng, standard_presets=standard)
        # show nonzero weights
        nz = df[df['weight'].fillna(0.0) > 0]
        print(' total_weight sum=', float(df['weight'].sum()))
        print(' nonzero MM count=', len(nz))
        print(' nonzero MM list (first 20):', nz['MM #'].astype(int).tolist()[:20])
        print(' unique distributions:', nz.get('distribution', pd.Series()).astype(str).str.lower().unique())
        # print first few rows
        print(nz[['MM #','weight','distribution','sigma_rad','sigma_azi','sigmax','sigmay']].head(12).to_string(index=False))

        # compute metrics
        metrics = plot_sum(df, normalize=True, fast=True, return_metrics_only=True)
        print(' metrics:', metrics)

if __name__ == '__main__':
    main()
