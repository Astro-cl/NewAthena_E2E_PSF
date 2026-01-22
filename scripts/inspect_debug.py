#!/usr/bin/env python3
import glob, os, json, sys
import numpy as np, pandas as pd
fig = 'Figures'
finals = sorted(glob.glob(os.path.join(fig, 'debug_combo_*_final_df.csv')), key=os.path.getmtime)
npzs = sorted(glob.glob(os.path.join(fig, 'debug_combo_*_Z_opt.npz')), key=os.path.getmtime)
metrics = sorted(glob.glob(os.path.join(fig, 'debug_combo_*_metrics.json')), key=os.path.getmtime)
rebuilds = sorted(glob.glob(os.path.join(fig, 'rebuild_result_*.csv')), key=os.path.getmtime)
out = {}
if not finals:
    print('ERROR: no debug_combo final_df found in Figures', file=sys.stderr)
    sys.exit(1)
fn = finals[-1]
out['final_df'] = fn
df = pd.read_csv(fn)
cols = ['weight', 'sigmax', 'sigmay', 'mux', 'muy']
stats = {}
for c in cols:
    if c in df.columns:
        arr = df[c].fillna(0).to_numpy(dtype=float)
        stats[c] = {'min': float(arr.min()), 'max': float(arr.max()), 'mean': float(arr.mean()), 'std': float(arr.std())}
out['final_stats'] = stats
if 'weight' in df.columns:
    top = df.sort_values('weight', ascending=False).head(20)
    out['top20'] = top[['weight', 'sigmax', 'sigmay', 'mux', 'muy']].to_dict(orient='records')
if npzs:
    zf = npzs[-1]
    data = np.load(zf)
    X = data['X'] if 'X' in data else (data['x'] if 'x' in data else None)
    Y = data['Y'] if 'Y' in data else (data['y'] if 'y' in data else None)
    Z = data['Z'] if 'Z' in data else (data['z'] if 'z' in data else None)
    out['Z_opt_file'] = zf
    if Z is not None:
        sZ = float(Z.sum())
        imax = int(Z.argmax())
        nrows, ncols = Z.shape
        imax_r, imax_c = divmod(imax, ncols)
        cx = float(X[imax_r, imax_c]) if X is not None else None
        cy = float(Y[imax_r, imax_c]) if Y is not None else None
        xs = float((X * Z).sum() / sZ) if X is not None else None
        ys = float((Y * Z).sum() / sZ) if Y is not None else None
        out['Z_stats'] = {'sum': sZ, 'max': float(Z.max()), 'peak_coord': (cx, cy), 'centroid': (xs, ys)}
if metrics:
    with open(metrics[-1]) as f:
        m = json.load(f)
    out['metrics_file'] = metrics[-1]
    out['metrics'] = m
if rebuilds:
    rb = rebuilds[-1]
    out['rebuild_result'] = rb
    rbd = pd.read_csv(rb)
    for c in ['sigmax', 'sigmay', 'mux', 'muy']:
        if c in rbd.columns:
            arr = rbd[c].fillna(0).to_numpy(dtype=float)
            out.setdefault('rebuild_stats', {})[c] = {'min': float(arr.min()), 'max': float(arr.max()), 'mean': float(arr.mean())}
print(json.dumps(out, indent=2))
