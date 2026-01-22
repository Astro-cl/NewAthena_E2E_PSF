from tools.run_sensitivity import build_aeff_mapping, load_mm_row_map, find_aeff_weights_for_choice, BASE_WORKBOOK, apply_mm_psf_choice_to_df
from main import load_gaussians_from_excel
import hashlib, numpy as np
am=build_aeff_mapping()
am['mm_row_map']=load_mm_row_map(BASE_WORKBOOK)
choice='1 keV [row6]'
map6=find_aeff_weights_for_choice(choice, am)
print('map6 sum', sum(map6.values()), 'nonzero', len([v for v in map6.values() if v>0]))
# build df
h=hashlib.sha1(repr({'A_eff':choice}).encode('utf8')).hexdigest()
seed=int(h[:16],16)%(2**32)
rng=np.random.default_rng(seed)
df=load_gaussians_from_excel(str(BASE_WORKBOOK), sheet='MM_PSF')
print('before df weight sum', float(df['weight'].sum()))
try:
    df['weight']=df['MM #'].astype(int).map(map6)
except Exception as e:
    print('map assignment failed', e)
print('after assign df weight sum', float(df['weight'].sum()))
choice_psf = '10% Variable Pseudo-Voigt 4.3" (alpha 10%)'
try:
    df2=apply_mm_psf_choice_to_df(df.copy(), choice_psf, am, rng)
    print('after apply_mm_psf_choice df weight sum', float(df2['weight'].sum()))
except Exception as e:
    print('apply failed', e)
