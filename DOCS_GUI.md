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
   Click `Load Excel File` → `Distributions/Test_Distribution.xlsx`

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

