import pandas as pd

# Load WITHOUT data_only=True (since pandas doesn't support that parameter)
aeff_raw = pd.read_excel('Distributions/test10kev.xlsx', sheet_name='A_eff', engine='openpyxl', header=None)
print(f"DF shape: {aeff_raw.shape}")

# Show all headers (row 0)
print("\n=== All Headers (Row 0) ===")
for i, header in enumerate(aeff_raw.iloc[0, :]):
    print(f"Col {i}: {header}")

# Show some sample data
print("\n=== Sample data from first 3 MM rows ===")
for i in range(3):
    print(f"Row {i+1} (MM {aeff_raw.iloc[i+1, 0]}): {aeff_raw.iloc[i+1, :].tolist()[:10]}")

# Look for "10 keV" columns
print("\n=== Looking for energy preset columns ===")
for i, header in enumerate(aeff_raw.iloc[0, :]):
    h_str = str(header).strip() if header else ""
    if "keV" in h_str or "keV" in h_str.lower():
        print(f"Col {i}: {h_str} -> data: {aeff_raw.iloc[1:5, i].tolist()}")

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
