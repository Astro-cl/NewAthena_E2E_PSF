import pandas as pd

fp = 'sensitivity/results/sensitivity_run_results.xlsx'
try:
    df = pd.read_excel(fp)
except Exception as e:
    print('ERROR_READING', e)
    raise

# helper to safely get column
def gcol(*names):
    for n in names:
        if n in df.columns:
            return n
    return None

# expected columns
hew_origin = gcol('hew_origin_arcsec','hew_origin')
hew_best = gcol('hew_best_arcsec','hew_best')
hew_opt = gcol('hew_opt_arcsec','hew_opt')
eef_origin = gcol('eef90_origin_arcsec','eef90_origin')
eef_best = gcol('eef90_best_arcsec','eef90_best')
eef_opt = gcol('eef90_opt_arcsec','eef90_opt')

print('RESULTS_FILE:', fp)
print('ROWS:', len(df))


def summarize(metric_name, orig_col, best_col, opt_col):
    if orig_col is None:
        print(f'No origin column for {metric_name} found; skipping')
        return
    print('\n---', metric_name, 'summary ---')
    if best_col and best_col in df.columns:
        df['_best_delta'] = df[best_col] - df[orig_col]
        bmean = df['_best_delta'].mean()
        bstd = df['_best_delta'].std()
        bmin = df['_best_delta'].min()
        bmax = df['_best_delta'].max()
        print(f'Best - Origin: mean={bmean:.6g}, std={bstd:.6g}, min={bmin:.6g}, max={bmax:.6g}')
        # top improvements (largest negative delta) and worsenings (largest positive)
        print('\nTop 5 improvements (best vs origin):')
        print(df.sort_values('_best_delta').head(5)[['_best_delta','input_file']].to_string(index=False))
        print('\nTop 5 worsenings (best vs origin):')
        print(df.sort_values('_best_delta', ascending=False).head(5)[['_best_delta','input_file']].to_string(index=False))
    else:
        print('No best column; skipping best-origin stats')
    if opt_col and opt_col in df.columns:
        df['_opt_delta'] = df[opt_col] - df[orig_col]
        omean = df['_opt_delta'].mean()
        ostd = df['_opt_delta'].std()
        omin = df['_opt_delta'].min()
        omax = df['_opt_delta'].max()
        print(f'Opt - Origin: mean={omean:.6g}, std={ostd:.6g}, min={omin:.6g}, max={omax:.6g}')
        print('\nTop 5 improvements (opt vs origin):')
        print(df.sort_values('_opt_delta').head(5)[['_opt_delta','input_file']].to_string(index=False))
        print('\nTop 5 worsenings (opt vs origin):')
        print(df.sort_values('_opt_delta', ascending=False).head(5)[['_opt_delta','input_file']].to_string(index=False))
    else:
        print('No opt column; skipping opt-origin stats')


summarize('HEW', hew_origin, hew_best, hew_opt)
summarize('EEF90', eef_origin, eef_best, eef_opt)

# cleanup
for c in ['_best_delta','_opt_delta']:
    if c in df.columns:
        del df[c]
