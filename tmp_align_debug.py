import pandas as pd
path='Distributions/NewTest_Distribution.xlsx'
align_df = pd.read_excel(path, sheet_name='Alignment', engine='openpyxl')
print('loaded cols len, rows len', len(align_df.columns), len(align_df))
print('col sample:', list(align_df.columns)[:12])
print('position-like?', any('position' in str(c).lower() or str(c).lower().strip().startswith('pos') for c in align_df.columns))
if any('position' in str(c).lower() or str(c).lower().strip().startswith('pos') for c in align_df.columns):
    tmp = align_df.copy()
    pos_col=None
    for c in tmp.columns:
        if 'position' in str(c).lower() or str(c).lower().strip().startswith('pos'):
            pos_col=c; break
    if pos_col is None:
        tmp['Position #'] = pd.to_numeric(tmp.index + 1, errors='coerce')
    else:
        tmp['Position #'] = pd.to_numeric(tmp[pos_col], errors='coerce')
    print('pos_col:',repr(pos_col),'tmp.shape',tmp.shape,'non-null',int(tmp['Position #'].notna().sum()))
    tmp = tmp[tmp['Position #'].notna()]
    head = tmp.head(10)
    for i, row in head.iterrows():
        print('row',i,'pos',row['Position #'])
        # print any d_align_z-like cols
        for c in row.index:
            cn = str(c).lower()
            if 'd_align' in cn and 'z' in cn:
                print('  col',c,'val',row.get(c))
    # show value of d_align_z via to_numeric on first row
    r = tmp.iloc[0]
    for c in r.index:
        cn = str(c).lower()
        v = r.get(c)
        if 'd_align' in cn and 'z' in cn:
            print('first row raw',c, v, type(v))
            num = pd.to_numeric(v, errors='coerce')
            print('first row numeric', num)
else:
    print('no position-like columns')
