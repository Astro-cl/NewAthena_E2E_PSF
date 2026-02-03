import pandas as pd
from pathlib import Path

wb = Path('Distributions/NewTest_Distribution.xlsx')
print('Workbook:', wb)
xl = pd.ExcelFile(wb)
print('Sheets:', xl.sheet_names)

if 'Alignment' in xl.sheet_names:
    df = xl.parse('Alignment', header=0)
    print('\nAlignment columns:')
    print(list(df.columns))
    print('\nFirst 60 rows (truncated):')
    try:
        print(df.head(60).to_string())
    except Exception:
        print(df.head(20))
    # locate possible d_align_z columns (case-insensitive)
    candidates = [c for c in df.columns if 'd_align' in str(c).lower() or 'align' in str(c).lower()]
    print('\nCandidate alignment-related columns:')
    for c in candidates:
        print('-', c)
    # show Position # column values if present
    pos_cols = [c for c in df.columns if 'position' in str(c).lower()]
    if pos_cols:
        pos_col = pos_cols[0]
        print('\nFound Position column:', pos_col)
        try:
            uniq = list(sorted(df[pos_col].dropna().unique()))
            print('Unique sample positions (first 40):', uniq[:40])
        except Exception:
            print('Could not extract unique positions for', pos_col)
    else:
        print('\nNo Position column detected in Alignment sheet.')
    # check for exact or variant d_align_z
    dz_cols = [c for c in df.columns if 'd_align_z' in str(c).lower()]
    print('\nExact d_align_z-like columns:')
    print(dz_cols)
    if dz_cols and pos_cols:
        print('\nSample values for', dz_cols[0])
        try:
            print(df[[pos_col, dz_cols[0]]].head(20).to_string())
        except Exception:
            print(df[[pos_col, dz_cols[0]]].head(20))
else:
    print('No Alignment sheet found.')

# search all sheets for headers containing d_align_z or d_align
print('\nSearching all sheets for headers with "d_align_z" or "d_align"...')
found = []
for s in xl.sheet_names:
    try:
        df2 = xl.parse(s, header=0)
        for c in df2.columns:
            if 'd_align_z' in str(c).lower() or 'd_align' in str(c).lower():
                found.append((s, c))
    except Exception:
        pass
print('Found occurrences:')
for s,c in found[:200]:
    print('-', s, ':', c)

print('\nFinished dump.')
