# Documentation Summary

> **Navigation:** [README.md](README.md) | [DOCS_GUI.md](DOCS_GUI.md) | [QUICKSTART.txt](QUICKSTART.txt)

This document provides a quick reference to the core modules and their purposes.

## Core Modules

| Module | Description |
|--------|-------------|
| [main.py](main.py) | Core loader, perturbation applier, projection helpers and PSF summation/plotting utilities |
| [gui_distributions.py](gui_distributions.py) | Tkinter-based GUI to load Excel workbooks, choose standard distributions, apply presets and export updated workbooks |
| [distributions_rotated.py](distributions_rotated.py) | Low-level rotated 2D distribution implementations (Gaussian, pseudo-Voigt) and helpers to evaluate PSF matrices |
| [optimize_mm_rows.py](optimize_mm_rows.py) | Utilities to reorder and optimize MM row layouts |
| `tools/` | Helper scripts and utilities (moved from top-level); legacy/debug copies retained under `scripts/legacy/` |

## Documentation Files

| Document | Purpose |
|----------|---------|
| [README.md](README.md) | Main project documentation with usage guide, GUI manual, and examples |
| [QUICKSTART.txt](QUICKSTART.txt) | Quick start instructions for common tasks |
| [DOCS_GUI.md](DOCS_GUI.md) | A_eff and vignetting export guide |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Guidelines for contributors |
| [RELEASE_NOTES.md](RELEASE_NOTES.md) | Release history and change log |
| [SENSITIVITY_QUICKSTART.txt](SENSITIVITY_QUICKSTART.txt) | Sensitivity pipeline guide |

## Recent Changes

- **(v9)** Interactive MM Selector viewer: `main.py` now opens a Tkinter
  window with a Row → Petal → MM checkbox tree and a live E2E PSF + EEF
  figure; `--export-package` bypasses the GUI for headless runs.
- **(v9)** `plot_sum()` gains `return_fig` and `figsize` parameters;
  subplot titles use `fig.text()` + `get_tightbbox` placement.
- Added module-level documentation to `main.py` and explanatory notes
- Added `CONTRIBUTING.md` and `DOCS_SUMMARY.md` to guide contributors
- Documented the new aggregated modified pseudo-Voigt fit and GUI EEF export behavior
- Documented HEW degradation plus energy-dependent sigma broadening behavior
- Cleaned up duplicate documentation between DOCS_GUI.md and README.md

## Modeling Notes

- HEW degradation broadening is applied per-position in quadrature:
	`sigma_broadened = sqrt(sigma_base^2 + sigma_extra^2)`
- Energy-dependent broadening then scales both `sigma_rad` and `sigma_azi` using the
	`MM HEW degradation energy` sheet (energy/factor table with linear interpolation):
	`sigma_final = sigma_broadened * f(E)`
- Final degraded/scaled sigma values are written to MM_PSF columns I/J and used by
	the aggregated PSF computation.
- See [README.md](README.md) section "Energy-Dependent Broadening (v8)" for full details.

## Recommended Next Steps

- Add per-function docstrings where missing in `distributions_rotated.py` and `optimize_mm_rows.py`
- Generate an API reference (e.g., using Sphinx) for public helpers
- Add examples/fixtures in `Distributions/` for common workflows

