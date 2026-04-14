GUI Guide — A_eff & Vignetting Export
=====================================

> **Note:** For general GUI usage with screenshots and step-by-step workflows, see the [GUI User Manual in README.md](README.md#gui-user-manual). This document focuses on A_eff handling and vignetting export workflows.

This guide explains the GUI behavior relevant to `A_eff` handling and
vignetting export-time copying.

Overview
--------
The GUI is implemented in `gui_distributions.py`:
```bash
python3 gui_distributions.py
```

Key tabs:
- `Load File` — open an Excel workbook containing MM configuration
- `MM Configuration` — select which mirror modules (MMs) to include
- `A_eff` — apply effective-area weights to selected MMs
- `Preview / Export` — inspect and export generated tables

A_eff Workflow
--------------
A_eff presets are read from the `A_eff` sheet (MM # in column A, weights in column B).

**Supported expression forms:**
- `J` — copy value from column letter `J`
- `gaussian(J, sigma)` — sample per-MM around column value
- `J+gaussian(0,20%*J)` — additive gaussian noise

**GUI options:**
- `Fixed Value` — apply single numeric A_eff to selected MMs
- `Standard Distribution` — evaluate preset per-MM

**Export:** A_eff values are written during the Export step.

Vignetting Export-Time Copy
---------------------------
The A_eff tab includes: `Apply vignetting factors when exporting`

When enabled with a standard A_eff preset, the export routine copies the chosen
preset column from `Vignetting rotazi` and `Vignetting rotrad` sheets into column B.

**Matching logic:**
1. Direct column letter (e.g., `K`) → use that column
2. Otherwise scan headers (lowercased, whitespace removed) for substring match

Example Workflow
----------------

1. Launch and load workbook:
   ```bash
   python3 gui_distributions.py
   ```
   Click `Load Excel File` → `Distributions/TestDistribution.xlsx`

2. Apply A_eff preset:
   - `MM Configuration`: Select MMs or use `Select All`
   - `A_eff`: Choose `Standard Distribution`, pick preset (e.g., `1 keV`)
   - Click `Apply to Selected MMs`

3. Enable vignetting copy and export:
   - Check `Apply vignetting factors when exporting`
   - `Preview / Export`: Choose `Save as new file`
   - Click `Export to Excel`

**Expected changes:**
- `A_eff` sheet: column A = `MM #`, column B = numeric values
- `Vignetting rotazi` / `Vignetting rotrad`: column B updated

Troubleshooting
---------------
- Workbook layout variations: GUI tolerates headers in different columns
- Missing A_eff values: Ensure preset or fixed value was applied
- Vignetting copy fails: Check preset string matches column letter or header

See Also
--------
- [README.md](README.md) — Main documentation with GUI screenshots
- [DOCS_SUMMARY.md](DOCS_SUMMARY.md) — Core module reference
- [CONTRIBUTING.md](CONTRIBUTING.md) — Contribution guidelines


## Release v5 (2026-02-07)

- GUI export: standard A_eff presets are evaluated per-MM and written as
  numeric values into column B of the `A_eff` sheet when exporting from
  the GUI. The GUI export clears column C to avoid writing adjusted A_eff.
- Percent-variable presets (e.g. "Variable 10% 1 keV") are synthesized
  into explicit gaussian expressions when loaded so sampling is deterministic
  and per-index.
- Preset energy tokens like `1 keV` are parsed and written into `C2` of
  vignetting sheets at export time.

## v6 (2026-04-01) — GUI polish and tools

 - Combobox click behavior (macOS): combobox fields now open on click and
   allow immediate mouse selection without requiring the Enter key. This
   improves UX on macOS trackpads and touch-based pointer devices.

- MM Configuration checkboxes: clicking the checkbox toggles only the
   checkbox state and does not inadvertently change selection focus —
   this makes selecting a single MM by clicking its checkbox more reliable.

 - A_eff caching script: the repository includes a helper script
    `tools/compute_aeff_values.py` which computes previously-formula-derived
    lookup columns and writes numeric A_eff values into the `A_eff` sheet
    so users can work with numeric presets even when Excel formulas aren't
    evaluated by the environment. Usage:

```bash
python3 tools/compute_aeff_values.py path/to/your_workbook.xlsx
```

   The script modifies the workbook in-place and is intended for
   post-processing workbooks generated with VLOOKUP-based A_eff presets.

See the `README.md` and `CHANGELOG.md` for release notes and details.

