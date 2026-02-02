import tempfile
import os
import pandas as pd
from main import load_gaussians_from_excel, load_aeff_weight_map_with_name, compute_total_rot_polar


def make_test_workbook(path):
    with pd.ExcelWriter(path, engine='openpyxl') as w:
        aeff = pd.DataFrame({'MM #': [1, 2], '0.25 keV': [10.0, 20.0]})
        aeff.to_excel(w, sheet_name='A_eff', index=False)
        mmpsf = pd.DataFrame({
            'MM #': [1, 2],
            'm_rad [arcsec]': [0.0, 0.0],
            'm_azi [arcsec]': [0.0, 0.0],
            'sigma_rad [arcsec]': [0.01, 0.01],
            'sigma_azi [arcsec]': [0.01, 0.01],
            'distribution': ['gaussian', 'gaussian']
        })
        mmpsf.to_excel(w, sheet_name='MM_PSF', index=False)
        mmcfg = pd.DataFrame({'MM #': [1, 2], 'Position #': [1, 2], 'x_MM [m]': [1.0, 0.0], 'y_MM [m]': [0.0, 1.0], 'r_MM [m]': [1.0, 1.0]})
        mmcfg.to_excel(w, sheet_name='MM configuration', index=False)
        align = pd.DataFrame({'Position #': [1, 2], 'd_align_rad [\u00b5m]': [0.0, 0.0], 'd_align_azi [\u00b5m]': [0.0, 0.0]})
        align.to_excel(w, sheet_name='Alignment', index=False)
        grav = pd.DataFrame({'Position #': [1, 2], 'd_grav_rotx [arcsec]': [0.0, 0.0], 'd_grav_roty [arcsec]': [0.0, 0.0]})
        grav.to_excel(w, sheet_name='Gravity offload', index=False)
        therm = pd.DataFrame({'Position #': [1, 2], 'd_therm_rotx [arcsec]': [0.0, 0.0], 'd_therm_roty [arcsec]': [0.0, 0.0]})
        therm.to_excel(w, sheet_name='Thermal', index=False)
        vazi = pd.DataFrame({'delta_arcsec': [-1.0, 0.0, 1.0], '1': [0.9, 0.9, 0.9], '2': [0.8, 0.8, 0.8]})
        vazi.to_excel(w, sheet_name='Vignetting rotazi', index=False)
        vrad = pd.DataFrame({'delta_arcsec': [-1.0, 0.0, 1.0], '1': [0.95, 0.95, 0.95], '2': [0.85, 0.85, 0.85]})
        vrad.to_excel(w, sheet_name='Vignetting rotrad', index=False)


if __name__ == '__main__':
    tf = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tf.close()
    path = tf.name
    try:
        make_test_workbook(path)
        print('Workbook:', path)
        mapping, colname = load_aeff_weight_map_with_name(path)
        print('A_eff mapping:', mapping, 'colname:', colname)
        df = load_gaussians_from_excel(path, sheet='MM_PSF')
        print('\nDataFrame attrs:', df.attrs)
        print('\nWeights:')
        print(df[['MM #','weight']])
        # compute rot maps directly
        mm_config_df = pd.read_excel(path, sheet_name='MM configuration', engine='openpyxl')
        mm_to_pos = dict(zip(mm_config_df['MM #'].astype(int), mm_config_df['Position #'].astype(int)))
        mmcfg_map = {int(r['MM #']): {'x_MM': r['x_MM [m]'], 'y_MM': r['y_MM [m]'], 'r_MM': r['r_MM [m]']} for _, r in mm_config_df.iterrows()}
        align_df = pd.read_excel(path, sheet_name='Alignment', engine='openpyxl')
        alignment_by_pos = {int(r['Position #']): {'d_align_rotazi': 0.0, 'd_align_rotrad': 0.0} for _, r in align_df.iterrows()}
        grav_df = pd.read_excel(path, sheet_name='Gravity offload', engine='openpyxl')
        thermal_df = pd.read_excel(path, sheet_name='Thermal', engine='openpyxl')
        _, _, rot_rad_map, rot_azi_map = compute_total_rot_polar(mm_to_pos, mmcfg_map, alignment_by_pos, {}, {})
        print('\nrot_rad_map:', rot_rad_map)
        print('rot_azi_map:', rot_azi_map)
        print('\nVignetting rotazi sheet:')
        print(pd.read_excel(path, sheet_name='Vignetting rotazi', engine='openpyxl'))
        print('\nVignetting rotrad sheet:')
        print(pd.read_excel(path, sheet_name='Vignetting rotrad', engine='openpyxl'))
    finally:
        try:
            os.remove(path)
        except Exception:
            pass

    # run alignment preset repro
    repro_alignment_preset()


def create_workbook_with_alignment_preset(path):
    import numpy as np
    mm_psf = pd.DataFrame({
        'MM #': [1],
        'm_rad [arcsec]': [0.0],
        'm_azi [arcsec]': [0.0],
        'sigma_rad [arcsec]': [4.3],
        'sigma_azi [arcsec]': [4.3],
    })
    aeff = pd.DataFrame({'MM #': [1], '0.25 keV': [1.0]})
    mm_conf = pd.DataFrame({
        'MM #': [1],
        'Position #': [1],
        'x_MM [m]': [1.0],
        'y_MM [m]': [0.0],
        'r_MM [m]': [1.0],
    })
    cols = []
    for i in range(15):
        cols.append(f'c{i}')
    align_row = [None] * 15
    align_row[0] = 1
    align_row[8] = 'preset_selected'
    align_row[12] = 2.0
    align_row[13] = -1.0
    align_df = pd.DataFrame([align_row], columns=cols)
    col_names = cols.copy()
    col_names[0] = 'Position #'
    align_df.columns = col_names
    grav = pd.DataFrame({'Position #': [1], 'd_grav_rotx [arcsec]': [0.0], 'd_grav_roty [arcsec]': [0.0]})
    therm = pd.DataFrame({'Position #': [1], 'd_therm_rotx [arcsec]': [0.0], 'd_therm_roty [arcsec]': [0.0]})
    vig_azi = pd.DataFrame({'delta_arcsec': [1.0, 2.0, 3.0], '1': [1.0, 1.2, 1.0]})
    vig_rad = pd.DataFrame({'delta_arcsec': [-2.0, -1.0, 0.0], '1': [0.9, 0.8, 1.0]})
    with pd.ExcelWriter(path, engine='openpyxl') as w:
        mm_psf.to_excel(w, sheet_name='MM_PSF', index=False)
        aeff.to_excel(w, sheet_name='A_eff', index=False)
        mm_conf.to_excel(w, sheet_name='MM configuration', index=False)
        align_df.to_excel(w, sheet_name='Alignment', index=False)
        grav.to_excel(w, sheet_name='Gravity offload', index=False)
        therm.to_excel(w, sheet_name='Thermal', index=False)
        vig_azi.to_excel(w, sheet_name='Vignetting rotazi', index=False)
        vig_rad.to_excel(w, sheet_name='Vignetting rotrad', index=False)


def repro_alignment_preset():
    tf = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tf.close()
    path = tf.name
    try:
        create_workbook_with_alignment_preset(path)
        print('\n--- alignment preset repro workbook:', path)
        df = load_gaussians_from_excel(path)
        print('df.attrs:', df.attrs)
        print('weights:')
        print(df[['MM #','weight']])
        mm_config_df = pd.read_excel(path, sheet_name='MM configuration', engine='openpyxl')
        mm_to_pos = dict(zip(mm_config_df['MM #'].astype(int), mm_config_df['Position #'].astype(int)))
        mmcfg_map = {int(r['MM #']): {'x_MM': r['x_MM [m]'], 'y_MM': r['y_MM [m]'], 'r_MM': r['r_MM [m]']} for _, r in mm_config_df.iterrows()}
        align_df = pd.read_excel(path, sheet_name='Alignment', engine='openpyxl')
        # Extract rotazi/rotrad from alignment dataframe via column indices 12 and 13 used in test
        rotazi = None
        rotrad = None
        try:
            rotazi = float(align_df.iloc[0, 12])
            rotrad = float(align_df.iloc[0, 13])
        except Exception:
            pass
        print('extracted rotazi, rotrad from Alignment row:', rotazi, rotrad)
    finally:
        try:
            os.remove(path)
        except Exception:
            pass
