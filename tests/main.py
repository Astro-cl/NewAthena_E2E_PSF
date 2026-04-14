import importlib.util
from pathlib import Path

# Shim that re-exports the repository root `main.py` so older tests
# which expect `tests/main.py` continue to work after test reorganization.
real_main = Path(__file__).resolve().parent.parent / 'main.py'
spec = importlib.util.spec_from_file_location('real_main', str(real_main))
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
# copy public attributes into this module's globals
for k, v in mod.__dict__.items():
    if not k.startswith('__'):
        globals()[k] = v
