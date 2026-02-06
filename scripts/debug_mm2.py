import tempfile, os, pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
from main import load_gaussians_from_excel, plot_sum


def _write_workbook(path, mm_rows, aeff_map, vig_factors_by_pos):
    with pd.ExcelWriter(path, engine='openpyxl') as w:
        pd.DataFrame(mm_rows).to_excel(w, sheet_name='MM_PSF', index=False)
        mm_config = []
        for mm in sorted([r['MM #'] for r in mm_rows]):
            if mm == 1:
                mm_config.append({'MM #': mm, 'x_MM [m]': 1.0, 'y_MM [m]': 0.0, 'r_MM [m]': 1.0})
            else:
                mm_config.append({'MM #': mm, 'x_MM [m]': 0.0, 'y_MM [m]': 1.0, 'r_MM [m]': 1.0})
        pd.DataFrame(mm_config).to_excel(w, sheet_name='MM configuration', index=False)
        aeff_df = pd.DataFrame({'MM #': list(aeff_map.keys()), 'weight': list(aeff_map.values())})
        aeff_df.to_excel(w, sheet_name='A_eff', index=False)
        align_rows = []
        for pos in sorted(vig_factors_by_pos.keys()):
            align_rows.append({'Position #': pos, 'd_align_rotazi [arcsec]': 0.0, 'd_align_rotrad [arcsec]': 0.0})
        pd.DataFrame(align_rows).to_excel(w, sheet_name='Alignment', index=False)
        deltas = [-1.0, 0.0, 1.0]
        vig_df = {'delta': deltas}
        for pos, factor in vig_factors_by_pos.items():
            vig_df[str(pos)] = [factor] * len(deltas)
        pd.DataFrame(vig_df).to_excel(w, sheet_name='Vignetting rotazi', index=False)
        pd.DataFrame(vig_df).to_excel(w, sheet_name='Vignetting rotrad', index=False)


if __name__ == '__main__':
    mm_rows = [
        {'MM #': 1, 'm_rad [arcsec]': 0.0, 'm_azi [arcsec]': 0.0, 'sigma_rad [arcsec]': 1.0, 'sigma_azi [arcsec]': 1.0},
        {'MM #': 2, 'm_rad [arcsec]': 10.0, 'm_azi [arcsec]': 0.0, 'sigma_rad [arcsec]': 1.0, 'sigma_azi [arcsec]': 1.0},
    ]
    with tempfile.TemporaryDirectory() as td:
        f = os.path.join(td, 'mm2_only.xlsx')
        _write_workbook(f, mm_rows, {1: 0.0, 2: 1.0}, {1: 1.0, 2: 1.0})
        df = load_gaussians_from_excel(f)
        print('=== DataFrame head ===')
        print(df.to_string())
        print('columns:', df.columns.tolist())
        print('weight sum:', float(df['weight'].sum()) if 'weight' in df.columns else 'no weight')
        try:
            m = plot_sum(df, return_metrics_only=True, fast=True, debug=True)
            print('metrics:', m)
        except Exception as e:
            print('plot_sum exception:', type(e), e)
