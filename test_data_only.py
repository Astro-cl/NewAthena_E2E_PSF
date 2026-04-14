import openpyxl
import pandas as pd

# Load with data_only=True using openpyxl directly
wb = openpyxl.load_workbook('Distributions/test10kev.xlsx', data_only=True)
ws = wb.active
print(f"Sheet name: {ws.title}")

# Read into DataFrame using openpyxl
data = []
for row in ws.iter_rows(values_only=True):
    data.append(row)

df = pd.DataFrame(data)
print(f"DF shape: {df.shape}")

# Show all headers (row 0)
print("\n=== All Headers (Row 0) ===")
for i, header in enumerate(df.iloc[0, :]):
    print(f"Col {i}: {header}")

# Show some sample data
print("\n=== Sample data from first 3 rows ===")
for i in range(3):
    print(f"Row {i+1}: {df.iloc[i+1, :].tolist()[:10]}")

# Look for "10 keV" columns
print("\n=== Looking for energy values in headers ===")
for i, header in enumerate(df.iloc[0, :]):
    h_str = str(header).strip() if header else ""
    if "keV" in h_str or "keV" in h_str.lower():
        print(f"Col {i}: {h_str} -> data: {df.iloc[1:5, i].tolist()}")
