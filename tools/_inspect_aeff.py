import pandas as pd
from pathlib import Path
p=Path('Distributions/Test_Distribution.xlsx')
print('file exists:',p.exists(),p)
raw=pd.read_excel(p,sheet_name='A_eff',engine='openpyxl',header=None)
print('shape',raw.shape)
header_row=None
for r in range(min(40,raw.shape[0])):
    v=raw.iloc[r,0]
    if isinstance(v,str) and v.strip().lower().replace(' ','') in {'mm#','mm'}:
        header_row=r
        break
print('header_row',header_row)
if header_row is None:
    print(raw.iloc[:10].to_string())
else:
    data=raw.iloc[header_row+1:]
    mm_vals=pd.to_numeric(data.iloc[:,0],errors='coerce')
    mm_nonnull=mm_vals.dropna().astype(int).tolist()
    print('MM count',len(mm_nonnull))
    col_names=raw.iloc[header_row,:].fillna('').astype(str).tolist()
    numeric_cols=[]
    for j in range(1,data.shape[1]):
        col=pd.to_numeric(data.iloc[:,j],errors='coerce')
        nonna=col.dropna()
        if len(nonna)>0:
            numeric_cols.append(col_names[j] if j<len(col_names) else f'col{j}')
    print('Numeric A_eff columns count',len(numeric_cols))
    print('Sample:',numeric_cols[:20])
