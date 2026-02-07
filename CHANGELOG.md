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
