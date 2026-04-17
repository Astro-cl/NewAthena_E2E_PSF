# Release Notes

This file contains a concise history of notable changes across releases.

## Release v8 (2026-04-17)

This release introduces four major features: off-axis pointing, defocus,
HEW degradation, and batch combination processing. It also includes
significant improvements to A_eff handling, sigma persistence, and
export packaging.

### Off-axis pointing

- Off-axis angle (user-specified in arcmin) is decomposed into X/Y rotation
  components via `offaxis × 60 / √2` (arcsec) and written to a new
  **"Extra PSF shifts"** sheet (columns B/C: `d_extra_rotx`, `d_extra_roty`).
- Previously off-axis was injected into Thermal columns; the new
  architecture cleanly separates the two perturbation sources.
- The extra rotations are added to the total rotation alongside alignment,
  gravity, and thermal contributions in `compute_total_rot_polar()`.

### Defocus

- Defocus (user-specified in mm) is converted to µm (`× 1e3`) and written
  to column D (`d_extra_z [µm]`) of the "Extra PSF shifts" sheet.
- The loader converts µm → m (`× 1e-6`) and adds `d_extra_z` to the total
  Z displacement (alignment + gravity + thermal + extra).
- Centroid shift is computed via Z-axis projection:
  `dm_x = d_z_total × x_MM / (12 − z_MM)`, where 12 m is the focal length.

### HEW degradation

- New sheet support: **"MM HEW degradation rotazi"** and
  **"MM HEW degradation rotrad"** with H–K lookup tables mapping
  (Row #, angle [arcmin], energy [keV]) → HEW degradation [arcsec].
- Per-position: angle from arcmin → arcsec, energy selection from C2,
  linear interpolation via `np.interp()`, result written to column B.
- Sigma broadening: `σ_new = √(σ_base² + (HEW / 2√(2·ln2))²)`.
  Rotrad → broadens `sigma_rad`; rotazi → broadens `sigma_azi`.
- Write-back:
  - Per-position `sigma_extra` → "Extra PSF degradations" sheet (cols B/C).
  - Final degraded sigma → MM_PSF columns I/J (`sigma_rad_deg`,
    `sigma_azi_deg` in arcsec).
- VLOOKUP resolver: when openpyxl strips cached formula values from MM_PSF
  D/E (base sigma), a Python fallback resolves the nested VLOOKUP chain
  (MM# → Row# via "MM configuration", Row# → sigma from preset table at
  M30:Q45). Resolved values are persisted as plain numbers.
- Preset distribution table shifted from column K to column M to avoid
  conflict with the new I/J columns.

### Batch combinations

- New `--batch-combinations <file.xlsx>` CLI argument for automated
  multi-configuration runs.
- Combinations file format: columns B (name), C (off-axis arcmin),
  D (energy keV), E (defocus mm), optional F (run_mode).
- Per-configuration: copies base workbook, writes "Extra PSF shifts" and
  energy values, runs the full analysis pipeline in headless mode
  (`matplotlib.use('Agg')`, noop `plt.show()`).
- Export packaging: per-config ZIP under `Exports/Export_<input>_<ts>/`
  containing workbook, FITS, PNGs, and fit parameters.
- Aggregated results workbook with per-config HEW, EEF, A_eff metrics.

### Other improvements

- Robust A_eff formula evaluation fallback for Excel lookup patterns.
- Dynamic `Aeff_loss` sum (no fixed 600-row assumption).
- Prefer packaged workbook for A_eff aggregation sums.
- Pearson4 fitting skipped in coarse/quick mode for performance.
- Default mode set to coarse; extra-fine mode removed.
- PSF PNG export now matches requested pixel resolution exactly.

### Test suite

- 71 tests passing (20 new integration tests):
  - 5 off-axis rotation tests
  - 6 defocus tests + 3 loader edge-case tests
  - 14 HEW degradation tests (interpolation + broadening)
  - 1 batch combinations end-to-end test

### Documentation

- New: [DOCS_FEATURES_APRIL2026.md](DOCS_FEATURES_APRIL2026.md) — extensive
  technical documentation of all v8 features with code references, formulas,
  data flow diagrams, and commit timeline.

## Release v7 (2026-04-14)

- Repository reorganization: tests grouped by concern under `tests/`
  (metrics, vignetting, io, gui, integration, plots, distributions).
- Utilities moved to `tools/`; agent scaffolding and debug scripts removed.
- Documentation and docstring pass across core modules.
- Fitting formulas documented (modified pseudo-Voigt, Pearson4, King).
- Export package behavior documented with verification steps.
- 42 tests passing after reorganization.

## Release v6 (2026-04-01)

- GUI polish: improved combobox click behavior on macOS; MM Configuration
  checkbox toggling improved.
- Added `tools/compute_aeff_values.py` for caching A_eff numeric columns.
- Repo hygiene: removed large artifacts from index; `.gitignore` updates.

## Release v5 (2026-02-07)

- GUI A_eff export semantics changed: when exporting from the GUI, a
  selected standard A_eff preset is evaluated per-MM and the resulting
  numeric A_eff values are written into column B of the `A_eff` sheet.
  The GUI export clears column C to avoid writing adjusted/vignetted A_eff
  values from the interactive export. The CLI (`main.py`) preserves the
  previous behavior and continues to populate column C when run
  non-interactively.
- Percent-variable presets are synthesized into explicit gaussian forms
  when loaded from the workbook (e.g. a preset named "Variable 10% 1 keV"
  with a Values cell of `L` becomes internally `gaussian(L,10%*L)`) so the
  evaluator can sample correctly.
- Preset-energy parsing improved: when a preset name contains an energy
  token such as `1 keV`, the GUI now parses that numeric energy and writes
  it into cell `C2` of the vignetting sheets at export time. The export
  routine prefers an explicit free-energy selection but falls back to the
  `keV` token if present.
- The preset evaluator (`_evaluate_aeff_preset_for_mm`) is used at export
  time so per-MM draws are consistent with the Apply action; a numeric
  fallback uses the `A_eff_base` value and the parsed percent when the
  direct evaluation fails.
- Replaced noisy debug prints with Python `logging` calls across the GUI
  module.
- Misc: repository cleanup (removed stale previews) and branch `cleanup/v5`
  merged into `main`.

## Release v4 (2026-02-03)

- Fixed dz→dm projection and ensured per‑MM dz outputs are correct when
  running `--log-dz` (proper polar/cartesian projection and use of
  alignment/gravity/thermal rotations).
- Re-ordered vignetting application to run after A_eff initialization and
  added bookkeeping fields `aeff_base`, `aeff_adjusted`, and
  `aeff_vig_factor` so plots and aggregations use the adjusted effective
  area.
- Improved vignetting parsing to support both two-column (delta→factor)
  sheets and per-position columns.
- GUI: added `Apply vignetting factors when exporting` checkbox in the
  `A_eff` tab; when enabled, the selected preset column from the
  vignetting sheets is copied into column B during export.

## Release v3 (2026-01-28)

- Repository cleanup and removal of legacy temporary tooling.
- Deterministic per-MM sampling for presets and parity between CSV/Excel
  exports.
- Added unit tests and pytest configuration for CI.
- Renamed `sensivitiy` → `sensitivity` and updated references.

## Release v2 (2025-12-15)

- Initial public release with core PSF generation, placement strategies,
  Excel I/O, GUI and CLI analysis utilities.


*For a detailed changelog or PR-level history, see the Git commit log.*
- GUI A_eff export: GUI-selected standard presets are evaluated per-MM and
  written numerically into column B of `A_eff`; GUI export clears column C.
- Percent-variable presets are synthesized to explicit gaussian forms
  (e.g. "Variable 10% 1 keV" → `gaussian(L,10%*L)`) to ensure correct
  deterministic per-index sampling.
- Preset energy tokens like `1 keV` are parsed and written into `C2` of
  vignetting sheets at export time.

