# CHANGELOG

This file summarizes notable changes across releases (human-readable).

## v5 — 2026-02-07
- GUI A_eff export: GUI-selected standard presets evaluated per-MM and
	numeric results written to column B; GUI export clears column C.
- Percent-variable presets synthesized to gaussian forms for deterministic
	sampling.
- Preset energy parsing: tokens like `1 keV` propagate to `C2` in
	vignetting sheets on export.

## v4 — 2026-02-03
- Vignetting reorder and bookkeeping (`aeff_base`, `aeff_adjusted`,
	`aeff_vig_factor`).
- GUI option to apply vignetting during export (copies selected vignette
	column into export).

## v3 — 2026-01-28
- Clean up repository for release: remove in-memory/pickle input flow.
- Add deterministic per-MM sampling and preset fuzzy-matching for reproducible inputs.
- Add unit tests and CI-friendly pytest configuration; tests pass locally.

## v2 — 2025-12-15
- Initial public release: core PSF generation, placement, GUI and CLI.

## v8 — 2026-04-17
- **Off-axis pointing:** decomposed into d_extra_rotx / d_extra_roty
	(arcmin × 60 / √2 → arcsec) and written to a dedicated "Extra PSF
	shifts" sheet instead of overwriting Thermal columns.
- **Defocus:** written as d_extra_z [µm] to the same "Extra PSF shifts"
	sheet (mm × 1e3 → µm); loader converts µm → m and projects via
	d_z × x_MM / (12 − z_MM) for centroid shifts.
- **HEW degradation:** new "MM HEW degradation rotazi" / "rotrad" sheets
	with H–K lookup tables; per-position interpolation of angle → HEW
	degradation; sigma broadening via √(σ_base² + (HEW/2√(2 ln 2))²).
- **Sigma writeback:** degraded sigma → MM_PSF I/J; sigma_extra → "Extra
	PSF degradations" B/C; VLOOKUP-resolved base sigma persisted to D/E
	as plain numbers.
- **Batch combinations:** `--batch-combinations` CLI for multi-config
	runs; ZIP packaging per config under Exports/Export_<input>_<ts>/;
	aggregated results workbook; headless plt.show noop.
- **Preset table shift:** MM_PSF preset distribution table moved from
	column K to column M to avoid I/J conflict.
- **A_eff improvements:** robust formula evaluation fallback, dynamic
	Aeff_loss sums, prefer packaged workbook for aggregation.
- **Performance:** Pearson4 skipped in coarse/quick mode; default mode
	set to coarse; extra-fine mode removed.
- Tests: 71 passed (20 new integration tests for off-axis, defocus, HEW
	degradation, and batch combinations).

## v7 — 2026-04-14
- Repository reorganization: tests grouped by concern; utilities moved
	to tools/; agent scaffolding removed.
- Documentation and docstring pass across core modules.
- Fitting formulas documented (pseudo-Voigt, Pearson4, King).
- Tests: 42 passed after reorganization.

## v6 — 2026-04-01
- GUI: improved combobox behaviour on macOS so clicking the field opens the
	dropdown and allows immediate mouse selection (no Enter required).
- GUI: MM Configuration checkbox click toggles selection without losing
	focus — single-MM selection and checkbox toggling are now responsive.
- Added `compute_aeff_values.py`: standalone script to compute and cache
	lookup-based A_eff numeric columns so workbooks no longer require formula
	evaluation to read A_eff presets.
- Logging: removed leftover debug prints in the GUI and normalized to
	`logging` usage for quieter runtime output.
- Repository: removed large example distributions/figures from the index and
	added `.gitignore` entries to avoid committing generated artifacts.
