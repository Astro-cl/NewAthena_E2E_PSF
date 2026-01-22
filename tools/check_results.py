import pandas as pd, os, sys
fn='Distributions/sensitivity_results.xlsx'
if not os.path.exists(fn):
    print('MISSING', fn)
    sys.exit(1)
df=pd.read_excel(fn, engine='openpyxl')
print('COLUMNS:', df.columns.tolist())
print(df.head(10).to_string(index=False))
