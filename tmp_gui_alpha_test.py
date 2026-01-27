import tkinter as tk
from gui_distributions import ExtendedGUI
import pandas as pd
import numpy as np
from tkinter import messagebox as _mbox

# Non-interactive dialogs
_mbox.askyesno = lambda *a, **k: True
_mbox.showinfo = lambda title, msg: print(f"INFO: {title}: {msg}")
_mbox.showwarning = lambda title, msg: print(f"WARN: {title}: {msg}")
_mbox.showerror = lambda title, msg: print(f"ERROR: {title}: {msg}")

root = tk.Tk()
root.withdraw()
app = ExtendedGUI(root)

# Prepare mm config
n = 600
cols = 15
# minimal mm_config_df to allow mapping
mm_cfg = pd.DataFrame({'MM #': list(range(1, n+1)), 'Position #': list(range(1, n+1))})
app.mm_config_df = mm_cfg

# Setup standard distributions
psf_name = '50% Variable Pseudo-Voigt 8" (alpha 10%)'
app.standard_distributions = {
    psf_name: {
        'name': psf_name,
        'type': 'pseudo-voigt',
        'sigma_rad': None,
        'sigma_azi': None,
        'alpha_rad': None,
        'alpha_azi': None,
    }
}
app.data_type_checkboxes['MM_PSF'].set(True)
app.update_data_type_tabs()
app.refresh_standard_distribution_controls()

# Select preset and apply
psf_combo = app.distribution_widgets['MM_PSF']['std_dist_combo']
psf_combo.set(psf_name)
app.on_mm_psf_standard_selected('MM_PSF')

# Prepare selection and generate
app.selected_mm_numbers = list(range(1, n+1))
app.generate_data('MM_PSF')

df = app.data_dfs.get('MM_PSF')
if df is None:
    print('No data generated')
else:
    print(df[['MM #','sigma_rad [arcsec]','sigma_azi [arcsec]','alpha_rad','alpha_azi']].head(20))

root.destroy()
