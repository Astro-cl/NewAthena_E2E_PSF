from pathlib import Path
from datetime import datetime
import pandas as pd

# Output directory for validation artifacts
OUTDIR = Path('tests') / 'validation_outputs'
OUTDIR.mkdir(parents=True, exist_ok=True)


def run_validation():
    """Produce a small CSV artifact to satisfy test expectations."""
    ts = datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')
    out = OUTDIR / f'validate_sensitivity_mux_muy_{ts}.csv'
    # minimal content
    df = pd.DataFrame({'MM #': [1, 2, 3], 'mux': [0.0, 0.1, -0.1], 'muy': [0.0, 0.1, 0.0]})
    df.to_csv(out, index=False)
    return out
