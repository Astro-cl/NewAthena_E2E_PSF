import pandas as pd
import pathlib
import math

root = pathlib.Path('Figures/diagnostics')
combos = sorted([p for p in root.iterdir() if p.is_dir() and p.name.startswith('combo_')])
summary = []
for c in combos:
    runf = c / 'df_run.csv'
    optf = c / 'df_opt.csv'
    if not runf.exists() or not optf.exists():
        continue
    dr = pd.read_csv(runf)
    do = pd.read_csv(optf)
    # align by MM #
    if 'MM #' in dr.columns and 'MM #' in do.columns:
        dr2 = dr.set_index('MM #')
        do2 = do.set_index('MM #')
        keys = sorted(set(dr2.index) & set(do2.index))
    else:
        dr2 = dr
        do2 = do
        keys = [i for i in dr2.index if i in do2.index]
    nrows = len(keys)
    diffs = { 'weight':0, 'mux':0, 'muy':0, 'sigmax':0, 'sigmay':0, 'theta_deg':0 }
    examples = {k:None for k in diffs}
    for k in keys:
        r = dr2.loc[k]
        o = do2.loc[k]
        for col in diffs:
            if col=='theta_deg':
                name = 'theta_degrees'
            else:
                name = col
            rv = r[name] if name in r.index else float('nan')
            ov = o[name] if name in o.index else float('nan')
            try:
                rvf = float(rv)
                ovf = float(ov)
                if math.isfinite(rvf) and math.isfinite(ovf):
                    tol = max(1e-8, abs(rvf)*1e-6, abs(ovf)*1e-6)
                    if abs(rvf-ovf) > tol:
                        diffs[col]+=1
                        if examples[col] is None:
                            examples[col] = (int(k), rvf, ovf)
            except Exception:
                pass
    summary.append((c.name, nrows, diffs, examples))

for s in summary:
    name, nrows, diffs, examples = s
    print(name, 'rows=', nrows)
    for k,v in diffs.items():
        print('  ', k, 'mismatches=', v, 'example=', examples[k])
    print()
