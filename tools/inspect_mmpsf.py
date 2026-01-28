import pickle, pandas as pd, json, sys
p='sensitivity/input/20260125T000456Z_2_A_eff6_keV_MM_PSF50_Variable_Pseudo-Voigt_8_alpha_10_AlignmentStandard_medialario_Gravity_offloadGZ_Thermal30_deg_FMS_Tilt.xlsx.pkl'
with open(p,'rb') as fh:
    sheets=pickle.load(fh)
mm = sheets.get('MM_PSF')
# canonical columns
srad_col = next((c for c in mm.columns if 'sigma_rad' in str(c).lower()), None)
sazi_col = next((c for c in mm.columns if 'sigma_azi' in str(c).lower()), None)
alpha_r = next((c for c in mm.columns if 'alpha_rad' in str(c).lower()), None)
alpha_a = next((c for c in mm.columns if 'alpha_azi' in str(c).lower()), None)

srad = pd.to_numeric(mm[srad_col], errors='coerce') if srad_col else pd.Series([], dtype=float)
sazi = pd.to_numeric(mm[sazi_col], errors='coerce') if sazi_col else pd.Series([], dtype=float)
alpha_r_s = pd.to_numeric(mm[alpha_r], errors='coerce') if alpha_r else pd.Series([], dtype=float)
alpha_a_s = pd.to_numeric(mm[alpha_a], errors='coerce') if alpha_a else pd.Series([], dtype=float)

def summarize(ser):
    ser_n = ser.dropna().astype(float)
    return {
        'count': int(ser_n.size),
        'mean': float(ser_n.mean()) if ser_n.size else None,
        'min': float(ser_n.min()) if ser_n.size else None,
        'max': float(ser_n.max()) if ser_n.size else None,
    }

out = {
    'srad_col': srad_col,
    'sazi_col': sazi_col,
    'alpha_rad_col': alpha_r,
    'alpha_azi_col': alpha_a,
    'sigma_rad_stats': summarize(srad),
    'sigma_azi_stats': summarize(sazi),
    'alpha_rad_stats': summarize(alpha_r_s),
    'alpha_azi_stats': summarize(alpha_a_s),
    'sigma_rad_first_100': [float(x) if pd.notna(x) else None for x in srad.iloc[:100].to_list()],
    'sigma_azi_first_100': [float(x) if pd.notna(x) else None for x in sazi.iloc[:100].to_list()],
    'alpha_rad_first_100': [float(x) if pd.notna(x) else None for x in alpha_r_s.iloc[:100].to_list()] if alpha_r else None,
    'alpha_azi_first_100': [float(x) if pd.notna(x) else None for x in alpha_a_s.iloc[:100].to_list()] if alpha_a else None,
}
print(json.dumps(out, indent=2))
