# Release Notes

This file contains a concise history of notable changes across releases.

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

