#!/usr/bin/env python3
import glob, os, sys, json
import numpy as np, pandas as pd
fig='Figures'
finals=sorted(glob.glob(os.path.join(fig,'debug_combo_*_final_df.csv')), key=os.path.getmtime)
inputs=sorted(glob.glob(os.path.join(fig,'debug_combo_*_input_df.csv')), key=os.path.getmtime)
out={}
if not finals:
    print('ERROR: no final_df', file=sys.stderr); sys.exit(1)
fn=finals[-1]
out['final_df']=fn
fdf=pd.read_csv(fn)
if not inputs:
    print(json.dumps({'error':'no input_df found'}))
    sys.exit(0)
inf=inputs[-1]
out['input_df']=inf
idf=pd.read_csv(inf)
# align by index if same length
if len(fdf)!=len(idf):
    out['note']='row counts differ: final %d vs input %d' % (len(fdf), len(idf))
# compare relevant cols
cols=['sigmax','sigmay','mux','muy','weight']
diffs={}
for c in cols:
    if c in fdf.columns and c in idf.columns:
        a=fdf[c].fillna(0).to_numpy(dtype=float)
        b=idf[c].fillna(0).to_numpy(dtype=float)
        # if lengths differ, compare min length
        L=min(len(a),len(b))
        d=a[:L]-b[:L]
        diffs[c]={'min':float(d.min()),'max':float(d.max()),'mean':float(d.mean()),'std':float(d.std())}
        # report large changes
        idxs=np.where(np.abs(d)>1e-9)[0]
        diffs[c]['n_changed']=int(len(idxs))
out['diffs']=diffs
print(json.dumps(out,indent=2))
