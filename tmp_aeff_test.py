import tkinter as tk
from gui_distributions import ExtendedGUI
import pandas as pd
from tkinter import messagebox as _mbox

# Non-interactive dialogs
_mbox.askyesno = lambda *a, **k: True
_mbox.showinfo = lambda title, msg: print(f"INFO: {title}: {msg}")
_mbox.showwarning = lambda title, msg: print(f"WARN: {title}: {msg}")
_mbox.showerror = lambda title, msg: print(f"ERROR: {title}: {msg}")

root = tk.Tk()
root.withdraw()
app = ExtendedGUI(root)

# Load MM config and A_eff sheet from known test file
path = 'Distributions/Test_Distribution.xlsx'
print('Using baseline:', path)
try:
    app.mm_config_df = pd.read_excel(path, sheet_name='MM configuration', engine='openpyxl')
    print('Loaded MM configuration rows:', len(app.mm_config_df))
except Exception as e:
    print('Could not load MM configuration:', e)

try:
    aeff_raw = pd.read_excel(path, sheet_name='A_eff', engine='openpyxl', header=None)
    app.aeff_raw_df = aeff_raw
    # Populate existing weights map
    app.aeff_weights = {}
    for r in range(1, aeff_raw.shape[0]):
        mm = aeff_raw.iloc[r, 0] if aeff_raw.shape[1] > 0 else None
        w = aeff_raw.iloc[r, 1] if aeff_raw.shape[1] > 1 else None
        if pd.isna(mm) or mm == '':
            continue
        try:
            mm_i = int(float(mm))
        except Exception:
            continue
        try:
            w_f = float(w)
        except Exception:
            continue
        app.aeff_weights[mm_i] = w_f
    # Load presets
    app.load_standard_aeff_presets(aeff_raw)
    print('Loaded A_eff presets:', list(app.aeff_standard_presets.keys())[:6])
except Exception as e:
    print('Could not load A_eff sheet:', e)

# Select some MMs
app.selected_mm_numbers = list(range(1, 11))
print('Selected MMs:', app.selected_mm_numbers)

# Choose a preset if available
if app.aeff_standard_presets:
    preset = list(app.aeff_standard_presets.keys())[0]
    app.aeff_selected_preset_var.set(preset)
    print('Using preset:', preset)
else:
    print('No A_eff presets found; will try fixed mode')
    app.aeff_mode_var.set('fixed')
    app.aeff_fixed_var.set('1.0')

# Call apply
app.apply_aeff_to_selected()

# Dump resulting weights for first 20 selected
for mm in app.selected_mm_numbers[:20]:
    print('MM', mm, 'weight:', app.aeff_weights.get(mm))

root.destroy()
