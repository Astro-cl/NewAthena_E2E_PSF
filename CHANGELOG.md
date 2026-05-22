# CHANGELOG

This file summarizes notable changes across releases (human-readable).

## v9.2 — 2026-05-22
- **HEW Contribution Ranking panel:** new bottom pane in the interactive MM
  Selector viewer. Vectorised leave-one-out analysis (O(N) over a shared polar
  grid) ranks all checked MMs by their individual HEW contribution. Results
  shown in a `ttk.Treeview` table (columns: Rank, MM #, ΔHEW ″, Row, Petal)
  coloured red/green for degrading/improving MMs. Multi-select via Shift-click
  (range) and Cmd/Ctrl-click (toggle on macOS); right-click → *Select in tree*
  syncs the MM tree checkboxes to the highlighted rows.
- **HEW Contribution Map:** *Map* button (enabled after each ranking) opens a
  `Toplevel` scatter plot of all MM physical positions (`x_MM [m]` / `y_MM [m]`
  from the MM configuration sheet, displayed in metres). Ranked MMs are drawn
  as colour-coded circles (RdYlGn colormap, symmetric zero-centred around
  ΔHEW = 0); unranked MMs shown as small black crosses. Axis limits are fixed
  to the full MM footprint regardless of the current selection.
- Test suite: 98 tests passing (0 failures).

## v9.1 — 2026-05-22
- **Vignetting bug fix (single / per_pos modes):** the row-by-row weight
  application loop and the final reconcile pass now handle all three vignetting
  table modes (`per_row_energy`, `per_pos`, `single`). Previously only
  `per_row_energy` was wired up, so `single` and `per_pos` sheets silently
  produced no effect on weights or `aeff_vig_factor` columns.
- **Vignetting abs() consistency:** every `np.interp()` call in the vignetting
  path now receives `abs(rotation_value)`. Vignetting tables are expected to
  use non-negative delta values; negative and positive rotations of the same
  magnitude now always yield the same factor.
- Test suite: 98 tests passing (0 failures).

## v9 — 2026-05-22
- **Interactive MM Selector viewer:** launching `main.py` without
	`--output` / `--export-package` now opens a Tkinter window with a
	collapsible Row → Petal → MM tree (checkbox-based) on the left and a
	live-updating E2E PSF + EEF figure on the right.
- Tree groups modules by Row and Petal (read from the "MM configuration"
	sheet); only MMs with `aeff_adjusted > 0` are listed.
- Select All / None buttons; partial-row selection indicated by ▣ symbol.
- "Update E2E & EEF" re-renders the figure for the checked subset.
- Window stays hidden until the initial render is complete.
- Batch / `--export-package` mode bypasses the GUI and calls `plot_sum`
	directly so no Tkinter window appears during automated runs.
- `plot_sum()` gains `return_fig` and `figsize` keyword arguments.
- Figure titles aligned via `fig.text()` + `draw_event` callback using
	`get_tightbbox` per subplot column so titles clear secondary axes.

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
- **Energy-dependent sigma scaling:** new "MM HEW degradation energy"
	sheet with Energy [keV] vs scaling factor table; linear interpolation
	at selected energy; multiplicative scaling of σ_rad and σ_azi applied
	after angle-based broadening (runs independently of rotazi/rotrad).
- **Sigma writeback:** degraded sigma → MM_PSF I/J; sigma_extra → "Extra
	PSF degradations" B/C; VLOOKUP-resolved base sigma persisted to D/E
	as plain numbers.
- **Batch combinations:** `--batch-combinations` CLI for multi-config
	runs; ZIP packaging per config under Exports/Export_<input>_<ts>/;
	aggregated results workbook; headless plt.show noop; parallel child
	subprocess execution via ThreadPoolExecutor.
- **Preset table shift:** MM_PSF preset distribution table moved from
	column K to column M to avoid I/J conflict.
- **A_eff improvements:** robust formula evaluation fallback, dynamic
	Aeff_loss sums, prefer packaged workbook for aggregation.
- **Performance:** Pearson4 skipped in coarse/quick mode; default mode
	set to coarse; extra-fine mode removed; batch formula scan skipped
	via BATCH_NO_FORMULAS env var; parallel child execution.
- Tests: 82 passed (31 new integration tests for off-axis, defocus, HEW
	degradation, energy scaling, and batch combinations).

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
