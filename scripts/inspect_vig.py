import pandas as pd
xls='Distributions/TestDistribution6_working.xlsx'
mmc = pd.read_excel(xls, sheet_name='MM configuration', engine='openpyxl')
print('MM configuration columns:', mmc.columns.tolist())
if 'Position #' in mmc.columns:
    rows = mmc.index[mmc['Position #']==1].tolist()
    cfg_row = rows[0]+1 if rows else 1
else:
    cfg_row = 1
print('cfg_row for position 1 (1-based data row):', cfg_row)
vdf = pd.read_excel(xls, sheet_name='Vignetting rotazi', engine='openpyxl', header=None)
print('VDF shape:', vdf.shape)
if vdf.shape[1]>7:
    matches = vdf[vdf.iloc[:,7]==cfg_row]
    print('rows matching cfg_row in col H indexes (0-based):', matches.index.tolist())
    for idx, r in matches.iterrows():
        i_val = r.iloc[8] if vdf.shape[1]>8 else None
        j_val = r.iloc[9] if vdf.shape[1]>9 else None
        k_val = r.iloc[10] if vdf.shape[1]>10 else None
        print('row', idx+1, 'I=', repr(i_val), 'J=', repr(j_val), 'K=', repr(k_val))
else:
    print('no col H in vdf')
try:
    print('cell K827 value repr:', repr(vdf.iat[826,10]))
except Exception as e:
    print('K827 error or not present:', e)
