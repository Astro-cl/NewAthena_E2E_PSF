#!/usr/bin/env python3
import subprocess, json, sys
from pathlib import Path

def run_per_mm(pickle_path):
    cmd = [sys.executable, 'tools/per_mm_contributions.py', pickle_path, '--n-r', '300', '--n-theta', '120', '--top', '100']
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(f"per_mm_contributions failed: {proc.stderr[:2000]}")
    return json.loads(proc.stdout)

def summarize(a, b):
    out = {}
    out['pickle_a'] = a['pickle']
    out['pickle_b'] = b['pickle']
    for k in ['r_50_arcsec','r_90_arcsec','total_energy']:
        out[f'a_{k}'] = a.get(k)
        out[f'b_{k}'] = b.get(k)
        try:
            out[f'delta_{k}'] = (b.get(k) - a.get(k)) if (a.get(k) is not None and b.get(k) is not None) else None
        except Exception:
            out[f'delta_{k}'] = None
    # top contributors MM# sets
    top_a = [t['MM #'] for t in a.get('top_contributors_90', [])]
    top_b = [t['MM #'] for t in b.get('top_contributors_90', [])]
    set_a = set(top_a)
    set_b = set(top_b)
    out['top90_a'] = top_a[:20]
    out['top90_b'] = top_b[:20]
    out['top90_jaccard'] = len(set_a & set_b) / len(set_a | set_b) if (set_a | set_b) else None
    # compute top differences by c90
    c90_map_a = {t['MM #']: t['c90'] for t in a.get('top_contributors_90', [])}
    c90_map_b = {t['MM #']: t['c90'] for t in b.get('top_contributors_90', [])}
    all_mms = sorted(set(list(c90_map_a.keys()) + list(c90_map_b.keys())))
    diffs = []
    for mm in all_mms:
        va = c90_map_a.get(mm, 0.0)
        vb = c90_map_b.get(mm, 0.0)
        diffs.append({'MM #': mm, 'a_c90': va, 'b_c90': vb, 'delta': vb - va})
    diffs_sorted = sorted(diffs, key=lambda x: abs(x['delta']), reverse=True)
    out['top_deltas'] = diffs_sorted[:20]
    return out

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: compare_pickles.py <pickleA> <pickleB>')
        sys.exit(2)
    a = run_per_mm(sys.argv[1])
    b = run_per_mm(sys.argv[2])
    s = summarize(a, b)
    print(json.dumps(s, indent=2))
