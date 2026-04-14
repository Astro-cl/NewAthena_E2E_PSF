import openpyxl as _openpyxl
import pandas as pd

# Load with data_only=True using openpyxl directly
wb = _openpyxl.load_workbook('Distributions/test10kev.xlsx', data_only=True)
ws = wb['A_eff']

print(f"Sheet: {ws.title}")

# Convert worksheet to DataFrame
data = []
for row in ws.iter_rows(values_only=True):
    data.append(row)

aeff_raw = pd.DataFrame(data)
print(f"DF shape: {aeff_raw.shape}")

# Show all headers (row 0)
print("\n=== All Headers (Row 0) ===")
for i, header in enumerate(aeff_raw.iloc[0, :]):
    print(f"Col {i}: {header}")

# Look for "10 keV" columns
print("\n=== Looking for energy preset columns with data ===")
for i, header in enumerate(aeff_raw.iloc[0, :]):
    h_str = str(header).strip() if header else ""
    if "keV" in h_str or "keV" in h_str.lower():
        data_vals = aeff_raw.iloc[1:5, i].tolist()
        print(f"Col {i}: {h_str} -> data: {data_vals}")

# Test the preset column matching logic
print("\n=== Testing preset column matching (like the GUI does) ===")
sel_preset_name = '10 keV'
preset_col = None
for j in range(min(aeff_raw.shape[1], 200)):
    try:
        header_val = str(aeff_raw.iloc[0, j]).strip()
        # Match: header contains the preset energy (e.g. '10 keV' in 'A_eff @10 keV')
        if str(sel_preset_name).strip() in header_val:
            preset_col = j
            print(f"Found preset column at {j}: {header_val}")
            break
    except Exception:
        continue

if preset_col is not None:
    print(f"\nValues in preset column for MM rows:")
    for r in range(1, 9):
        mm = aeff_raw.iloc[r, 0]
        val = aeff_raw.iloc[r, preset_col]
        print(f"  MM {mm}: {val} (type: {type(val).__name__})")
else:
    print("Could not find preset column for '10 keV'")
