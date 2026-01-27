from pathlib import Path
import re
import pandas as pd

inp = Path('sensivitiy/input')
baseline = Path('Distributions/Test_Distribution.xlsx')

csvs = sorted(inp.glob('*_MM_PSF_sampling_detailed_*.csv'))
if not csvs:
    print('No sampling CSVs found in', inp)
    raise SystemExit(0)

rows = []
for csv in csvs:
    df = pd.read_csv(csv)
    def find_col(key):
        for c in df.columns:
            cl = c.lower()
            if cl.startswith(f'sampled_{key}') or ('sampled' in cl and key in cl):
                return c
        for c in df.columns:
            cl = c.lower()
            if key in cl and ('raw' in cl or 'sampled' in cl):
                return c
        return None
    stats = {'file': csv.name}
    for k in ['sr','sa','ar','aa']:
        c = find_col(k)
        if c:
            from pathlib import Path
            import re
            import pandas as pd

            inp = Path('sensivitiy/input')
            baseline = Path('Distributions/Test_Distribution.xlsx')

            csvs = sorted(inp.glob('*_MM_PSF_sampling_detailed_*.csv'))
            if not csvs:
                print('No sampling CSVs found in', inp)
                raise SystemExit(0)


            def parse_gamma_cell(cell):
                if cell is None:
                    return None
                s = str(cell).strip()
                # gamma(mean, 50%*mean) or gamma(mean,50%*mean)
                m = re.search(r'gamma\s*\(?\s*([0-9.]+)\s*,\s*([0-9]+)%', s, flags=re.I)
                if m:
                    mean = float(m.group(1))
                    pct = float(m.group(2)) / 100.0
                    sigma = mean * pct
                    return mean, sigma
                # alternate pattern '50%*4.1303' or '50% * 4.1303'
                m2 = re.search(r'([0-9]+)%\s*\*\s*([0-9.]+)', s)
                if m2:
                    pct = float(m2.group(1)) / 100.0
                    mean = float(m2.group(2))
                    sigma = mean * pct
                    return mean, sigma
                # plain numeric
                m3 = re.search(r'^\s*([0-9.]+)\s*$', s)
                if m3:
                    mean = float(m3.group(1))
                    return mean, None
                return None


            def find_target_for_filename(baseline_df, fname):
                # search all string cells in baseline for a candidate preset name that appears in filename
                for _, row in baseline_df.iterrows():
                    for cell in row:
                        if isinstance(cell, str) and cell.strip():
                            token = cell.strip()
                            if token in fname or token.replace(' ', '_') in fname:
                                prow = row
                                parsed = []
                                for cell2 in list(prow.iloc[10:20]):
                                    if pd.isna(cell2):
                                        parsed.append(None)
                                    else:
                                        parsed.append(parse_gamma_cell(cell2))
                                return parsed[:4]
                return [None, None, None, None]


            baseline_df = None
            try:
                baseline_df = pd.read_excel(baseline, sheet_name='MM_PSF', header=None)
            except Exception:
                baseline_df = None

            rows = []
            for csv in csvs:
                df = pd.read_csv(csv)

                def find_col(key):
                    for c in df.columns:
                        cl = c.lower()
                        if cl.startswith(f'sampled_{key}') or ('sampled' in cl and key in cl):
                            return c
                    for c in df.columns:
                        cl = c.lower()
                        if key in cl and ('raw' in cl or 'sampled' in cl):
                            return c
                    return None

                stats = {'file': csv.name}
                for k in ['sr', 'sa', 'ar', 'aa']:
                    c = find_col(k)
                    if c:
                        arr = pd.to_numeric(df[c], errors='coerce').dropna()
                        stats[f'{k}_n'] = int(arr.size)
                        stats[f'{k}_emp_mean'] = float(arr.mean())
                        stats[f'{k}_emp_std'] = float(arr.std(ddof=0))
                    else:
                        stats[f'{k}_n'] = 0
                        stats[f'{k}_emp_mean'] = None
                        stats[f'{k}_emp_std'] = None

                tgt = find_target_for_filename(baseline_df, csv.name) if baseline_df is not None else [None, None, None, None]
                for i, key in enumerate(['sr', 'sa', 'ar', 'aa']):
                    tv = tgt[i]
                    if tv is None:
                        stats[f'{key}_tgt_mean'] = None
                        stats[f'{key}_tgt_std'] = None
                        stats[f'{key}_mean_rel_err_pct'] = None
                        stats[f'{key}_std_rel_err_pct'] = None
                    else:
                        mean, std = tv
                        stats[f'{key}_tgt_mean'] = mean
                        stats[f'{key}_tgt_std'] = std
                        emp_mean = stats.get(f'{key}_emp_mean')
                        emp_std = stats.get(f'{key}_emp_std')
                        if emp_mean is not None and mean not in (None, 0):
                            stats[f'{key}_mean_rel_err_pct'] = 100.0 * (emp_mean - mean) / mean
                        else:
                            stats[f'{key}_mean_rel_err_pct'] = None
                        if emp_std is not None and std not in (None, 0):
                            stats[f'{key}_std_rel_err_pct'] = 100.0 * (emp_std - std) / std
                        else:
                            stats[f'{key}_std_rel_err_pct'] = None

                rows.append(stats)

            rep = pd.DataFrame(rows)
            outp = inp / 'validation_report.csv'
            rep.to_csv(outp, index=False)
            print('Wrote', outp)
            print(rep.head(20).to_string(index=False))
