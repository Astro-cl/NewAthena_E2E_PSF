import tkinter as tk
import traceback, sys, time
from pathlib import Path
import tempfile

# Stub messagebox and filedialog to avoid GUI modal blocks
import tkinter.messagebox as mb
import tkinter.filedialog as fd

mb.showinfo = lambda *a, **k: print('MSG showinfo', a)
mb.showwarning = lambda *a, **k: print('MSG showwarning', a)
mb.showerror = lambda *a, **k: print('MSG showerror', a)
mb.askyesno = lambda *a, **k: True
fd.asksaveasfilename = lambda *a, **k: str(Path(tempfile.gettempdir()) / 'export_test.xlsx')

import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parents[1]))
try:
    from gui_distributions import ExtendedGUI
except Exception:
    traceback.print_exc()
    sys.exit(1)

root = tk.Tk(); root.withdraw()
app = ExtendedGUI(root)
# Set minimal state to trigger A_eff pending export path
app.excel_path = str(Path('Distributions/Test_Distribution.xlsx'))
app.aeff_pending_export = True
app.selected_mm_numbers = [1,2,3]
app.enabled_data_types = []
# Use 'current' to save back to the loaded workbook
app.export_mode_var.set('current')
app.export_path_var.set('')

print('Calling export_to_excel...')
start = time.time()
try:
    app.export_to_excel()
    print('export_to_excel returned')
except Exception:
    print('Exception in export_to_excel:')
    traceback.print_exc()
finally:
    print('Duration', time.time()-start)
    root.destroy()
