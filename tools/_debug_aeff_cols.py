import pandas as pd
from pathlib import Path
p=Path('Distributions/Test_Distribution.xlsx')
raw=pd.read_excel(p,sheet_name='A_eff',engine='openpyxl',header=None)
# find header
header_row=0
for r in range(min(40,raw.shape[0])):
    v=raw.iloc[r,0]
    if isinstance(v,str) and v.strip().lower().replace(' ','') in {'mm#','mm'}:
        header_row=r
        break
print('header_row',header_row)
cols = raw.iloc[header_row,:].fillna('').astype(str).tolist()
data = raw.iloc[header_row+1:]
for j in range(1, data.shape[1]):
    name = cols[j] if j < len(cols) else f'col{j}'
    col = pd.to_numeric(data.iloc[:, j], errors='coerce')
    nonna = col.dropna()
    ssum = 0.0
    sabs = 0.0
    if nonna.size>0:
        ssum = float(nonna.sum())
        sabs = float(nonna.abs().sum())
    print(j, repr(name), 'nonna', len(nonna), 'sum', ssum, 'abs_sum', sabs)
