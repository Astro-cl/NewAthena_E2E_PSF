Documentation Summary
=====================

This repository contains the following core modules and short descriptions:

- `main.py`: Core loader, perturbation applier, projection helpers and PSF
  summation/plotting utilities.
- `gui_distributions.py`: Tkinter-based GUI to load Excel workbooks, choose
  standard distributions, apply presets and export updated workbooks.
- `distributions_rotated.py`: Low-level rotated 2D distribution implementations
  (Gaussian, pseudo-Voigt) and helpers to evaluate PSF matrices.
- `optimize_mm_rows.py`: Utilities to reorder and optimize MM row layouts.
- `scripts/`: Helper scripts for export and diagnostics (some debug helpers
  were removed during cleanup).

Recent documentation changes (2026-02-03):
- Added module-level documentation to `main.py` and small explanatory notes.
- Added `CONTRIBUTING.md` and this `DOCS_SUMMARY.md` to guide contributors.

Recommended next steps for fuller documentation:
- Add per-function docstrings where missing in `distributions_rotated.py` and
  `optimize_mm_rows.py`.
- Generate an API reference (e.g., using Sphinx) for public helpers.
- Add examples/fixtures in `Distributions/` for common workflows.
