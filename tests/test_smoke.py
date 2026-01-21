import subprocess
import sys
import importlib


def test_imports():
    # Ensure core modules import without raising
    importlib.import_module('gui_distributions')
    importlib.import_module('distributions_rotated')
    importlib.import_module('optimize_mm_rows')
    importlib.import_module('main')


def test_cli_help():
    # Running `python3 main.py --help` should exit with code 0 and print usage
    res = subprocess.run([sys.executable, 'main.py', '--help'], capture_output=True, text=True)
    assert res.returncode == 0
    assert 'usage' in res.stdout.lower() or 'usage' in res.stderr.lower()
