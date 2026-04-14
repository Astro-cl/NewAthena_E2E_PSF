import importlib.util
from pathlib import Path

# Shim to re-export the repository `sensitivity/sensitivity_run.py` module
# so tests that expect it under `tests/` continue to work after reorganization.
real_mod = Path(__file__).resolve().parent.parent / 'sensitivity' / 'sensitivity_run.py'
spec = importlib.util.spec_from_file_location('real_sensitivity', str(real_mod))
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
for k, v in mod.__dict__.items():
    if not k.startswith('__'):
        globals()[k] = v
