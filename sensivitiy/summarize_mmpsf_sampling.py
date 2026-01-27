#!/usr/bin/env python3
"""Summarize MM_PSF sampling CSVs: per-file stats and histograms."""
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
import numpy as np

IN = Path(__file__).resolve().parent / 'input'
OUT = Path(__file__).resolve().parent / 'figures'
OUT.mkdir(parents=True, exist_ok=True)

rows = []
for fp in sorted(IN.glob('*_MM_PSF_sampling_detailed_*.csv')):
    try:
        df = pd.read_csv(fp)
    except Exception:
        continue
    if 'sampled_sr_raw' not in df.columns:
        continue
    sr = pd.to_numeric(df['sampled_sr_raw'], errors='coerce').dropna()
    sa = pd.to_numeric(df['sampled_sa_raw'], errors='coerce').dropna() if 'sampled_sa_raw' in df.columns else pd.Series(dtype=float)
    r = {
        'file': fp.name,
        'n': int(len(df)),
        'sr_mean': float(sr.mean()) if len(sr) else None,
        'sr_std': float(sr.std()) if len(sr) else None,
        'sa_mean': float(sa.mean()) if len(sa) else None,
        'sa_std': float(sa.std()) if len(sa) else None,
    }
    rows.append(r)
    # histograms
    try:
        fig, ax = plt.subplots(1, 2, figsize=(10, 4))
        ax[0].hist(sr.values, bins=50)
        ax[0].set_title('sampled_sr_raw')
        ax[1].hist(sa.values, bins=50)
        ax[1].set_title('sampled_sa_raw')
        fig.suptitle(fp.name)
        fig.savefig(OUT / (fp.stem + '_hist.png'))
        plt.close(fig)
    except Exception:
        pass

if rows:
    pd.DataFrame(rows).to_csv(OUT / 'MM_PSF_sampling_summary.csv', index=False)
    print(f'Wrote summary for {len(rows)} files to {OUT}/MM_PSF_sampling_summary.csv')
else:
    print('No sampling CSVs found.')
