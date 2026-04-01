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
