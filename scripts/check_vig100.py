import pandas as pd
import numpy as np
wb='Distributions/NewTest_Distribution.xlsx'
df = pd.read_excel(wb, sheet_name='Vignetting rotazi', engine='openpyxl')
cols = list(df.columns)
print('columns:', cols)
if len(cols) < 2:
    print('sheet has fewer than 2 columns; head:')
    print(df.head(20))
else:
    a = pd.to_numeric(df[cols[0]], errors='coerce')
    b = pd.to_numeric(df[cols[1]], errors='coerce')
    matches = df[~a.isna() & (np.isclose(a, 100))]
    print('matches count:', len(matches))
    if len(matches) > 0:
        print(matches[[cols[0], cols[1]]].to_string(index=False))
    else:
        idx = np.argsort(a.fillna(np.inf))
        a_sorted = a.fillna(np.inf).to_numpy()[idx]
        b_sorted = b.fillna(np.nan).to_numpy()[idx]
        ins = np.searchsorted(a_sorted, 100)
        start = max(0, ins-3)
        end = min(len(a_sorted), ins+3)
        print('nearby values (sorted):')
        for i in range(start, end):
            print(i, a_sorted[i], b_sorted[i])
        ok = (~np.isnan(a_sorted)) & (~np.isnan(b_sorted)) & np.isfinite(a_sorted)
        if ok.sum() > 1:
            xs = a_sorted[ok]
            ys = b_sorted[ok]
            if not np.all(np.diff(xs) > 0):
                idx2 = np.argsort(xs)
                xs = xs[idx2]; ys = ys[idx2]
            val = float(np.interp(100, xs, ys, left=ys[0], right=ys[-1]))
            print('interp at 100:', val)
        else:
            print('not enough numeric points to interpolate')
