import tkinter as tk
from gui_distributions import ExtendedGUI
import pandas as pd
import numpy as np

root = tk.Tk()
root.withdraw()
app = ExtendedGUI(root)

# Prepare A_eff sheet and preset
# Build aeff_raw_df with columns A..O (0..14), MM numbers in col 0, values in col 1 (B)
df = pd.DataFrame(np.zeros((20,15)))
for i in range(20):
    df.iloc[i,0] = i
    df.iloc[i,1] = 100.0 + i
app.aeff_raw_df = df
# Add preset
preset_name = 'Variable 20% 7 keV'
app.aeff_standard_presets = {preset_name: 'B+gaussian(0,20%*B)'}
app.refresh_aeff_preset_controls()
# Simulate user selecting the preset
app.aeff_selected_preset_var.set(preset_name)
app.on_aeff_standard_selected()
print('A_eff expr label:', app.aeff_expr_label.cget('text'))

# Select some MMs and apply preset
app.selected_mm_numbers = [2, 5, 10]
app.apply_aeff_to_selected()
print('A_eff weights for selected MMs:')
for mm in app.selected_mm_numbers:
    print(mm, app.aeff_weights.get(mm))

# Now create MM_PSF tab widgets by enabling MM_PSF and building tabs
# Inject a minimal standard distribution definition with a safe key BEFORE building UI
app.standard_distributions = {
    'PSF_VAR': {
        'name': 'PSF_VAR',
        'type': 'gaussian',
        'sigma_rad': {'dist': 'gaussian', 'mean': 5.0, 'sigma': 1.0},
        'sigma_azi': {'dist': 'gaussian', 'mean': 6.0, 'sigma': 1.5},
        'alpha_rad': {'dist': 'fixed', 'value': 0.1},
        'alpha_azi': {'dist': 'fixed', 'value': 0.1},
    }
}
app.data_type_checkboxes['MM_PSF'].set(True)
app.update_data_type_tabs()
app.refresh_standard_distribution_controls()
# Simulate selecting the MM_PSF standard preset
std_combo = app.distribution_widgets['MM_PSF']['std_dist_combo']
std_combo.set('PSF_VAR')
# Call handler
app.on_mm_psf_standard_selected('MM_PSF')
print('MM_PSF std combo selection handled successfully.')

root.destroy()
