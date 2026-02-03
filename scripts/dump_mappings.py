import pandas as pd
from pathlib import Path

wb = Path('Distributions/NewTest_Distribution.xlsx')
print('Workbook:', wb)
xl = pd.ExcelFile(wb)
print('Sheets:', xl.sheet_names)

if 'MM configuration' in xl.sheet_names:
    mmc = xl.parse('MM configuration', header=0)
    print('\nMM configuration columns:')
    print(list(mmc.columns))
    # show first 40 rows
    print('\nFirst 40 rows:')
    try:
        print(mmc.head(40).to_string())
    except Exception:
        print(mmc.head(20))
    # build mm_to_pos and mm_config_map
    mm_to_pos = {}
    mm_config_map = {}
    if 'MM #' in mmc.columns:
        for i, row in mmc.iterrows():
            mmn = row.get('MM #')
            if pd.isna(mmn):
                continue
            try:
                mmn_i = int(pd.to_numeric(mmn, errors='coerce'))
            except Exception:
                continue
            pos = None
            if 'Position #' in mmc.columns and not pd.isna(row.get('Position #')):
                try:
                    pos = int(pd.to_numeric(row.get('Position #'), errors='coerce'))
                except Exception:
                    pos = None
            if pos is None:
                pos = len(mm_to_pos) + 1
            mm_to_pos[mmn_i] = pos
            x = row.get('x_MM [m]') if 'x_MM [m]' in mmc.columns else row.get('x_MM') if 'x_MM' in mmc.columns else None
            y = row.get('y_MM [m]') if 'y_MM [m]' in mmc.columns else row.get('y_MM') if 'y_MM' in mmc.columns else None
            r = row.get('r_MM [m]') if 'r_MM [m]' in mmc.columns else row.get('r_MM') if 'r_MM' in mmc.columns else None
            mm_config_map[mmn_i] = {
                'x_MM': float(pd.to_numeric(x, errors='coerce') or 0.0),
                'y_MM': float(pd.to_numeric(y, errors='coerce') or 0.0),
                'r_MM': float(pd.to_numeric(r, errors='coerce') or 0.0),
            }
    print('\nSample mm_to_pos (first 60 items):')
    for k in sorted(list(mm_to_pos.keys())[:60]):
        print(k, '->', mm_to_pos[k])
    print('\nSample mm_config_map (first 60 items):')
    for k in sorted(list(mm_config_map.keys())[:60]):
        print(k, '=>', mm_config_map[k])
else:
    print('No MM configuration sheet found.')

print('\nFinished mapping dump.')
