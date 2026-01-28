import re
from pathlib import Path
import pandas as pd

INPUT_DIR = Path('sensitivity/input')
if not INPUT_DIR.exists():
    print('No input dir:', INPUT_DIR)
    raise SystemExit(1)

files = sorted(INPUT_DIR.glob('*.xlsx'))
if not files:
    print('No xlsx files in', INPUT_DIR)
    raise SystemExit(0)

reports = []
for fp in files:
    try:
        df = pd.read_excel(fp, sheet_name='MM_PSF', engine='openpyxl')
    except Exception:
        # try headerless
        try:
            raw = pd.read_excel(fp, sheet_name='MM_PSF', engine='openpyxl', header=None)
            # try to promote a header row if appears
            first = raw.iloc[0].astype(str).str.strip().tolist()
            if any('m_rad' in str(x).lower() for x in first):
                raw.columns = first
                df = raw.iloc[1:].reset_index(drop=True)
            else:
                df = raw
        except Exception as e:
            reports.append((fp, 'error_reading_sheet', str(e)))
            continue

    # canonical column names
    cols = [c for c in df.columns]
    # try to find sigma columns
    sr = None
    sa = None
    for c in cols:
        lc = str(c).lower()
        if 'sigma' in lc and 'rad' in lc:
            sr = c
        if 'sigma' in lc and 'azi' in lc:
            sa = c
    if sr is None and sa is None:
        reports.append((fp, 'no_sigma_cols', None))
        continue

    # inspect first 26 data rows
    n_check = min(26, len(df))
    head = df.iloc[:n_check]

    # find template rows in sheet that indicate Fixed preset
    # search for any cell containing word 'fixed' (case-insensitive)
    expected = None
    for idx, row in df.iterrows():
        for cell in row:
            if isinstance(cell, str) and re.search(r'fixed', cell, flags=re.IGNORECASE):
                # try to read sigma values in this row (if columns exist)
                try:
                    v1 = None if sr is None else pd.to_numeric(df.at[idx, sr], errors='coerce')
                    v2 = None if sa is None else pd.to_numeric(df.at[idx, sa], errors='coerce')
                    if pd.notna(v1) or pd.notna(v2):
                        expected = (v1 if pd.notna(v1) else None, v2 if pd.notna(v2) else None)
                        break
                except Exception:
                    pass
        if expected is not None:
            break

    # if expected found, compare head values
    mismatches = []
    for i in range(n_check):
        r = head.iloc[i]
        v1 = None if sr is None else pd.to_numeric(r.get(sr), errors='coerce')
        v2 = None if sa is None else pd.to_numeric(r.get(sa), errors='coerce')
        if expected is not None:
            exp1, exp2 = expected
            # compare if expected is numeric
            if exp1 is not None and not pd.isna(exp1):
                if pd.isna(v1) or abs(v1 - exp1) > 1e-9:
                    mismatches.append((i+1, sr, v1, exp1))
            if exp2 is not None and not pd.isna(exp2):
                if pd.isna(v2) or abs(v2 - exp2) > 1e-9:
                    mismatches.append((i+1, sa, v2, exp2))
        else:
            # If no expected, still check for identical values across head? skip
            pass

    if expected is not None and mismatches:
        reports.append((fp, 'mismatch_first26', expected, mismatches[:10]))

# Print summary
if not reports:
    print('No issues found in', len(files), 'workbooks')
else:
    print('Found issues in', len(reports), 'workbooks:')
    for r in reports:
        print(r)
