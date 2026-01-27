import tkinter as tk
from gui_distributions import ExtendedGUI
import pandas as pd
import tkinter as tk
from gui_distributions import ExtendedGUI
import pandas as pd
import numpy as np
from tkinter import messagebox as _mbox

# Non-interactive test: override messagebox dialogs to avoid GUI prompts during automated smoke test
_mbox.askyesno = lambda *a, **k: True
_mbox.showinfo = lambda title, msg: print(f"INFO: {title}: {msg}")
_mbox.showwarning = lambda title, msg: print(f"WARN: {title}: {msg}")
_mbox.showerror = lambda title, msg: print(f"ERROR: {title}: {msg}")

root = tk.Tk()
root.withdraw()
app = ExtendedGUI(root)

# Prepare A_eff sheet with 600 MMs
n = 600
cols = 15
df = pd.DataFrame(np.zeros((n, cols)))
for i in range(n):
    df.iloc[i, 0] = i + 1  # MM # starting at 1
    df.iloc[i, 1] = 100.0 + i
app.aeff_raw_df = df

# Add preset for A_eff and refresh
import tkinter as tk
from gui_distributions import ExtendedGUI
import pandas as pd
import numpy as np
from tkinter import messagebox as _mbox

# Non-interactive test: override messagebox dialogs to avoid GUI prompts during automated smoke test
_mbox.askyesno = lambda *a, **k: True
_mbox.showinfo = lambda title, msg: print(f"INFO: {title}: {msg}")
_mbox.showwarning = lambda title, msg: print(f"WARN: {title}: {msg}")
_mbox.showerror = lambda title, msg: print(f"ERROR: {title}: {msg}")

root = tk.Tk()
root.withdraw()
app = ExtendedGUI(root)

# Prepare A_eff sheet with 600 MMs
n = 600
cols = 15
df = pd.DataFrame(np.zeros((n, cols)))
for i in range(n):
    df.iloc[i, 0] = i + 1  # MM # starting at 1
    df.iloc[i, 1] = 100.0 + i
app.aeff_raw_df = df

# Add preset for A_eff and refresh
preset_name = 'Variable 20% 7 keV'
app.aeff_standard_presets = {preset_name: 'B+gaussian(0,20%*B)'}
app.refresh_aeff_preset_controls()
app.aeff_selected_preset_var.set(preset_name)
app.on_aeff_standard_selected()

# Select all MMs and apply preset
app.selected_mm_numbers = list(range(1, n + 1))
app.apply_aeff_to_selected()
print('A_eff weights count:', len(app.aeff_weights))
print("A_eff tab status:", app.tab_status_labels.get('A_eff').cget('text') if app.tab_status_labels.get('A_eff') else 'N/A')

# Prepare standard distributions and enable MM_PSF/thermal/gravity tabs
# MM_PSF: create a pseudo-voigt Variable preset with no explicit alpha entries
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

# Add Thermal and Gravity presets so defaults are picked
app.thermal_standard_presets = {'0 deg FMS tilt': {'d_therm_x [µm]': 'fixed(0)', 'd_therm_y [µm]': 'fixed(0)', 'd_therm_z [µm]': 'fixed(0)', 'd_therm_rotz [arcsec]': 'fixed(0)'}}
app.gravity_standard_presets = {'GZ': {'d_grav_x [µm]': 'fixed(0)', 'd_grav_y [µm]': 'fixed(0)', 'd_grav_z [µm]': 'fixed(0)', 'd_grav_rotz [arcsec]': 'fixed(0)'}}

# Enable data tabs and build UI
app.data_type_checkboxes['MM_PSF'].set(True)
app.data_type_checkboxes['Thermal'].set(True)
app.data_type_checkboxes['Gravity offload'].set(True)
app.update_data_type_tabs()

# Refresh controls so defaults are applied
app.refresh_standard_distribution_controls()

# Check Thermal and Gravity defaults
therm_combo = app.distribution_widgets.get('Thermal', {}).get('therm_std_combo')
grav_combo = app.distribution_widgets.get('Gravity offload', {}).get('grav_std_combo')
print('Thermal combo selection:', therm_combo.get() if therm_combo else 'N/A')
print('Gravity combo selection:', grav_combo.get() if grav_combo else 'N/A')

# Now select the PSF preset and apply
psf_combo = app.distribution_widgets['MM_PSF']['std_dist_combo']
psf_combo.set(psf_name)
app.on_mm_psf_standard_selected('MM_PSF')

# Inspect alpha fields for MM_PSF
alpha_info = {}
if 'MM_PSF' in app.alpha_entries_by_type:
    for ui_param, widgets in app.alpha_entries_by_type['MM_PSF'].items():
        _, dist_box, entry_a, entry_b, _ = widgets
        alpha_info[ui_param] = {
            'dist_box': dist_box.get(),
            'value': entry_a.get(),
            'state': entry_a.cget('state') if hasattr(entry_a, 'cget') else 'unknown'
        }
print('MM_PSF alpha fields after applying standard pseudo-voigt preset:')
print(alpha_info)

# Check PSF tab status label
print('PSF tab status:', app.tab_status_labels.get('MM_PSF').cget('text') if app.tab_status_labels.get('MM_PSF') else 'N/A')

root.destroy()
