# NewAthenaE2EPSF_v3 - PSF Analysis Toolkit

[![Python CI](https://github.com/Astro-cl/NewAthena_E2E_PSF/actions/workflows/python-ci.yml/badge.svg)](https://github.com/Astro-cl/NewAthena_E2E_PSF/actions/workflows/python-ci.yml)

A comprehensive toolkit for PSF (Point Spread Function) modeling and analysis of mirror module configurations with support for multiple distribution types, perturbation analysis, and an interactive GUI.

> **Documentation, comments and Unit tests written by AI.**

## Documentation Index

| Document | Description |
|----------|-------------|
| [README.md](README.md) | This file - main project overview and guide |
| [QUICKSTART.txt](QUICKSTART.txt) | Quick start instructions for common tasks |
| [DOCS_GUI.md](DOCS_GUI.md) | GUI user manual with screenshots and workflows |
| [DOCS_SUMMARY.md](DOCS_SUMMARY.md) | Summary of core modules and API reference |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Guidelines for contributors |
| [RELEASE_NOTES.md](RELEASE_NOTES.md) | Release history and change log |
| [DOCS_FEATURES_APRIL2026.md](DOCS_FEATURES_APRIL2026.md) | v8 feature documentation (off-axis, defocus, HEW degradation, batch) |
| [SENSITIVITY_QUICKSTART.txt](SENSITIVITY_QUICKSTART.txt) | Sensitivity pipeline guide |

## Quick Links
- [Features](#features)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage](#usage)
- [Distribution Types](#distribution-types)

## Sensitivity Pipeline

The repository includes a small sensitivity-run template and driver under the `sensitivity/` folder.

- **Per-combo input workbooks:** generated Excel inputs are written to `sensitivity/input/` as `TIMESTAMP_index_<combo>.xlsx`. These workbooks are preserved by default but the runner will prune the folder to keep only the newest 100 files to avoid unbounded growth.
- **MM_PSF edits:** when the pipeline writes or expands the `MM_PSF` sheet it only modifies per‑MM input columns `B`..`H` (left-side per‑MM table). Columns to the right (template/tail columns) are preserved or restored from the baseline workbook.
- **Alpha masking:** for presets that are not pseudo‑voigt (i.e. Gaussian/Uniform), the `alpha_rad` and `alpha_azi` values in the `MM_PSF` per‑MM rows are set to `-` (columns G and H) to explicitly mark them as not applicable.
- **Alignment / Thermal / Gravity zeroing:** when a combo requests `Alignment=0`, the runner sets `d_align_rotazi` and `d_align_rotrad` to `0` in the `Alignment` sheet (only within columns B..G). Similarly, `Thermal=0` zeros `d_therm_rotx` and `d_therm_roty` in the `Thermal` sheet, and `Gravity offload=0` zeros `d_grav_rotx` and `d_grav_roty` in the `Gravity offload` sheet; other columns are left untouched.
- **Partial results:** as jobs complete the runner appends a per-job summary row to `sensitivity/results/sensitivity_run_partial.csv` (best-effort). The final consolidated Excel results are written at the end to `sensitivity/results/sensitivity_run_results.xlsx`.
- **Run modes:**
   - Generate-only: `python3 sensitivity/sensitivity_run.py --generate-only --baseline <file>` creates the per-combo input workbooks without executing jobs.
   - Full run: `python3 sensitivity/sensitivity_run.py --baseline <file>` generates inputs and runs the jobs, producing both the partial CSV as jobs finish and the final Excel at completion.


## Features

### Core Capabilities
- **Multiple Distribution Types**: Support for Gaussian and Pseudo-Voigt (mixture of Gaussian and Lorentzian) distributions with independent azimuthal and radial alpha parameters
- **Standard PSF Presets**: Load predefined MM_PSF distributions from Excel table with automatic parameter derivation for variable presets
- **Row-Wise MM Optimizer**: Optimize MM# assignments within rows to minimize HEW while keeping physical locations fixed
- **Rotation-Invariant HEW**: Polar grid integration eliminates orientation bias
- **Perturbation Analysis**: Alignment errors, gravity offload effects, and thermal deformations
- **Fast Mode**: Optimized computation with configurable sampling density
- **Multi-MM Analysis**: Process multiple mirror modules with different distributions
- **Excel Integration**: Load/save configurations with full parameter support and formula preservation
- **Interactive GUI**: User-friendly interface for parameter generation and configuration with standard/free preset modes
- **Figure Export**: High-resolution PNG exports of PSF and encircled energy plots

### Distribution Features
- **Fixed Values**: All mirror modules use identical parameters
- **Gaussian Distribution**: Parameters vary across MMs using normal distribution N(μ, σ²)
- **Uniform Distribution**: Parameters vary across MMs using uniform distribution U(min, max)
- **Alpha Parameter Control**: For Pseudo-Voigt, alpha_rad and alpha_azi can be fixed or vary per MM (automatically clamped to [0, 1])
- **Standard Presets**: Load predefined PSF distributions from the Excel preset table (MM_PSF sheet, starting at cell K1) with automatic variable parameter derivation
- **Free Mode**: Full manual control over all PSF parameters

## Installation
## Author

- **Ivo Ferreira** — primary author and maintainer.
- **Affiliation:** European Space Agency
- **Contact:** ivo.ferreira@esa.int

- **ORCID:** https://orcid.org/0000-0002-9501-862X



## Release History

### Release v2 (2025-12-15)

- Initial public release of the NewAthenaE2EPSF toolkit with core PSF
   generation, placement strategies and Excel I/O. Included Gaussian and
   Pseudo-Voigt support, basic GUI, and command-line analysis utilities.

### Release v3 (2026-01-28)

- Repository cleanup: removed legacy in-memory pickle flows and temporary
   debug tools.
- Deterministic per-MM sampling for presets and CSV/Excel parity in
   generation; added unit tests and pytest configuration to streamline CI.
- Renamed `sensivitiy` → `sensitivity` and updated Quickstart and tests.

### Release v4 (2026-02-03)

- Fixed dz→dm projection and ensured per‑MM dz outputs are correct when
   running `--log-dz` (proper polar/cartesian projection and use of
   alignment/gravity/thermal rotations).
- Re-ordered vignetting application to run after A_eff initialization and
   added explicit bookkeeping: `aeff_base`, `aeff_adjusted`, and
   `aeff_vig_factor` so plots and aggregations use the adjusted effective
   area.
- Improved vignetting parsing to support both two-column (delta→factor)
   sheets and per-position columns; MM300 now receives the combined
   vignette factor when multiple factors apply.
- GUI: added `Apply vignetting factors when exporting` checkbox in the
   `A_eff` tab; when enabled, the selected preset column from the
   vignetting sheets is copied into column B during export (matches by
   explicit column letter or header substring).

### Release v5 (2026-02-07)

- GUI A_eff export semantics changed: when exporting from the GUI, a
   selected *standard* A_eff preset is evaluated per-MM and the resulting
   numeric A_eff values are written into column B of the `A_eff` sheet.
   The GUI export path explicitly clears column C so adjusted/vignetted
   A_eff values are not written by the interactive export. The CLI
   (`main.py`) preserves the previous behavior and continues to populate
   column C with adjusted A_eff when run non-interactively.
- Percent-variable presets are synthesized into explicit gaussian forms
   when loaded from the workbook (e.g. a preset named "Variable 10% 1 keV"
   with a Values cell of `L` becomes internally `gaussian(L,10%*L)`) so the
   evaluator can sample correctly.
- Preset-energy parsing improved: when a preset name contains an energy
   token such as `1 keV`, the GUI now parses that numeric energy and writes
   it into cell `C2` of the vignetting sheets at export time. The export
   routine prefers an explicit free-energy selection (combobox) but falls
   back to the `keV` token in the preset name if present.
- The preset evaluator (`_evaluate_aeff_preset_for_mm`) is used at export
   time so per-MM draws are consistent with the Apply action; a numeric
   fallback uses the `A_eff_base` value and the parsed percent when the
   direct evaluation fails.
- Removed noisy debug prints and replaced them with Python `logging` calls
   across the GUI module so runtime logs are quieter in normal operation.
- Misc: small repository cleanup (removed stale preview workbooks under
   `Distributions/` and created a zipped backup). The branch `cleanup/v5`
   containing these updates was merged into `main` on 2026-02-07.

### Release v6 (2026-04-01)

- GUI polish: improved combobox click behavior on macOS so clicking a
   combobox opens the list and allows immediate mouse selection without
   requiring Enter; MM Configuration checkboxes now toggle reliably when
   clicked, improving single-MM selection.
 - Tooling: added `tools/compute_aeff_values.py` — a helper script to compute
    and cache A_eff lookup/formula results into numeric columns so workbooks
    can be used without Excel formula evaluation. See `tools/compute_aeff_values.py`
    for usage.
- Repo hygiene: removed large generated distribution/figure files from the
   repository index and added `.gitignore` entries to avoid accidental
   commits of generated artifacts.

### Release v8 (2026-04-17)

Major feature release introducing off-axis pointing, defocus, HEW
degradation, and batch combination processing.

- **Off-axis pointing**: decomposed into X/Y rotation components
   (`offaxis × 60 / √2` arcsec) and written to a dedicated "Extra PSF
   shifts" sheet, cleanly separated from thermal perturbations. Applied
   additively in `compute_total_rot_polar()`.
- **Defocus**: written as `d_extra_z [µm]` to the same "Extra PSF shifts"
   sheet (mm × 1e3 → µm); projected to centroid shifts via
   `d_z × x_MM / (12 − z_MM)` where 12 m is the focal length.
- **HEW degradation**: new "MM HEW degradation rotazi"/"rotrad" sheets
   with lookup tables (Row #, angle, energy → HEW arcsec); per-position
   interpolation; sigma broadening via
   `σ_new = √(σ_base² + (HEW / 2√(2·ln2))²)`. Results written to
   "Extra PSF degradations" (sigma_extra) and MM_PSF columns I/J
   (degraded sigma).
- **VLOOKUP resolver**: Python-based fallback resolves MM_PSF D/E formula
   values when openpyxl strips cached values; persists base sigma as plain
   numbers for round-trip stability.
- **Batch combinations**: `--batch-combinations` CLI for automated
   multi-configuration runs (off-axis, energy, defocus per row); per-config
   ZIP packaging under `Exports/`; aggregated results workbook; headless
   operation (no GUI blocking).
- **A_eff improvements**: robust formula evaluation fallback; dynamic
   `Aeff_loss` sums; prefer packaged workbook for aggregation.
- **Performance**: Pearson4 skipped in coarse/quick mode; default mode set
   to coarse; extra-fine mode removed.
- **Preset table shift**: MM_PSF distribution table moved from column K to
   column M to avoid conflict with new I/J degraded sigma columns.
- Tests: 71 passed (20 new integration tests for off-axis, defocus, HEW
   degradation, and batch combinations).
- Documentation: [DOCS_FEATURES_APRIL2026.md](DOCS_FEATURES_APRIL2026.md)
   provides extensive technical documentation of all v8 features.

### Release v7 (2026-04-14)

This release collects the repository reorganization, test refactor, and
documentation improvements implemented since **Release v6**. The
section below documents the representative commits, the exact fitting
formulas used by the codebase, the export options exposed to users, and
recommended release actions.

Representative commit summary (since v6)

- Remove agent scaffolding and debug helpers: deleted the M365 copilot
   scaffold and transient top-level debug scripts introduced by earlier
   automated tooling.
- Move utilities to `tools/`: relocated helper scripts (`compute_aeff_values.py`,
   `inspect_openpyxl.py`, `check_lookup_table.py`, `metrics_numba.py`, and
   friends) into `tools/` and preserved archival copies under
   `scripts/legacy/`.
- Tests reorganization: grouped tests by concern under `tests/` folders
   (metrics, vignetting, io, gui, integration, plots, distributions) and
   updated imports to use package-style imports (e.g., `tools.*`). The
   local test run reported 42 passed, 4 warnings (curve-fit related).
- Documentation updates: `README.md`, `DOCS_GUI.md` and
   `DOCS_SUMMARY.md` updated to reflect layout changes, updated usage
   examples, and new references to `tools/compute_aeff_values.py`.
- Module docstrings: added top-level documentation strings to
   `main.py` and `optimize_mm_rows.py` to clarify public APIs and
   expected behaviors for headless/CI use.

Fitting formulas and parameter descriptions

The following summarizes the analytic forms implemented in
`distributions_rotated.py` and the parameters exposed in the public
APIs.

1) Rotated 2D Gaussian (`gaussian_2d_rotated`)

    Let (x,y) be evaluation coordinates and (μ_x, μ_y) the center. For
    principal standard deviations σ_x, σ_y and rotation angle θ (radians
    unless `degrees=True`):

       dx = x - μ_x
       dy = y - μ_y
       th = θ
       c = cos(th); s = sin(th)
       invsx2 = 1/σ_x^2; invsy2 = 1/σ_y^2
       a = c^2*invsx2 + s^2*invsy2
       b = s*c*(invsx2 - invsy2)
       ccoef = s^2*invsx2 + c^2*invsy2

    Exponent: E = -0.5 * (a*dx^2 + 2*b*dx*dy + ccoef*dy^2)

    Output: amplitude * coeff * exp(E), where coeff = 1/(2π σ_x σ_y)
    if `normalize=True`, else coeff = 1.

    Parameters:
    - `mux`, `muy`: center (meters)
    - `sigmax`, `sigmay`: principal sigmas (meters)
    - `theta`: rotation angle (radians or degrees)
    - `amplitude`: multiplicative amplitude
    - `normalize`: produce unit-integral PDF before amplitude

2) Fitting models: Modified Pseudo-Voigt, Pearson Type IV, King

    The codebase supports three primary fitting profiles used across
    analysis and plotting: a modified separable Pseudo‑Voigt, Pearson
    Type IV (Pearson4), and a King/Moffat-style core+wing model. These
    are implemented to support analytic evaluation on rotated grids and
    to be compatible with discrete PSF matrix resampling.

    A) Modified separable Pseudo-Voigt (azimuthal × radial)

       - Each axis uses a 1D pseudo-Voigt: PV(u; σ, α) = (1 - α)*G(u; σ) + α*L(u; σ)
          where G(u; σ) = (1/√(2π)) exp(-u^2/2) and L(u; σ) = (1/π) * 1/(1+u^2).
       - The 2D shape is formed by PV_azi(azi_rot/σ_azi) * PV_rad(rad_rot/σ_rad)
          evaluated in the rotated principal-axis frame (consistent rotation
          convention with the Gaussian implementation).
       - `alpha` may be specified per-axis (`alphaazi`, `alpharad`) or a
          single `eta` mixing parameter may be used as fallback.
       - `normalize=True` divides by (σ_azi * σ_rad) to account for the
          normalization change of variables.

    B) Pearson Type IV (Pearson4)

       - The 1D Pearson4 unnormalized profile used in fits is:

             P(u) ∝ (1 + (u/σ)^2)^{-m} * exp(ν * atan(u/σ)),  with u = x - μ

       - Parameters: `μ` (center), `σ` (scale), `m` (tail-shape > 0), and
          `ν` (skew/asymmetry). ν = 0 gives a symmetric tail.
       - In 2D the implementation uses separable azimuthal/radial forms or
          a radial-only fit depending on the target data; normalization is
          computed numerically when required.

    C) King / Moffat-like profile

       - Radial form: K(r) = A * (1 + (r/α)^2)^{-β}, where r is radial
          distance, `α` is core scale and `β` controls wing steepness.
       - Parameters: `A` (amplitude), `α` (core scale), `β` (wing exponent).
       - Supports rotationally symmetric 2D evaluation; normalization
          computed analytically when possible or numerically otherwise.

3) Rotation and units

      - Rotation conventions are consistent across `gaussian_2d_rotated`,
          `pseudo_voigt_2d_rotated` and `eval_psf_matrix_rotated` so that
          analytic and discrete PSF evaluations align.
      - Public APIs use meters; the loader converts arcsec using the
          historical factor 1 arcsec = 12 * π / 180 / 3600 meters. Tests
          and callers must either provide meters or rely on loader
          conversions.

Additional fitted models

4) Pearson Type IV (Pearson4)

      The repository includes an implementation of a Pearson Type IV
      fit used to capture peaked, asymmetric shapes with heavy tails. The
      Pearson4 form used in fit helpers follows a numerically-stable
      parameterization where the unnormalized 1D profile is:

         P(u) ∝ (1 + (u/σ)^2)^{-m} * exp(ν * atan(u/σ))

      with u = x - μ. Parameters:
      - `μ`: center (meters)
      - `σ`: scale (meters)
      - `m`: tail-shape (>0)
      - `ν`: skew/asymmetry (0 = symmetric)

      The 2D usage in the code is either separable (azimuthal × radial)
      or radial-only depending on the fit target; normalization constants
      are computed numerically when required.

5) King profile (Moffat/King-like core+wings)

      A King/Moffat-like profile is available for fits that require a
      flat core with power-law wings. The standard radial form is:

         K(r) = A * (1 + (r/α)^2)^{-β}

      where r is radial distance, `α` is core scale and `β` the wing
      exponent. Parameters:
      - `A`: amplitude (flux)
      - `α`: core scale (meters)
      - `β`: wing exponent (>1)

      The implementation supports radial fits and rotationally symmetric
      2D evaluation; normalization is optional and computed analytically
      or numerically as appropriate.

Export behavior and available formats

GUI (`gui_distributions.py`):
- `Preview / Export` → `Save as new file` evaluates standard A_eff
   presets per‑MM and writes numeric A_eff to column B of the `A_eff`
   sheet; column C is cleared for adjusted values. When
   `Apply vignetting factors when exporting` is enabled the selected
   vignette column is copied into the exported workbook. Parsed preset
   energy tokens (e.g., `1 keV`) are written into vignetting `C2`.
- Plot context-menu exports:
   - PSF PNG (high-resolution)
   - Encircled Energy (EEF) PNG
   - EEF CSV (written to `CustomPSFs/` as `E2E_EEF_YYYYMMDD_HHMMSS.csv`)
   - Fit parameters CSV (aggregated fit coefficients)
   - FITS output for PSF matrices where applicable

CLI (`main.py`):
- Legacy behavior preserved: adjusted/derived A_eff values are recorded
   in column C by default for headless CLI runs. Use `--output` to save
   figures and `--no-gui` for headless execution. CLI accepts Excel and
   CSV inputs (including the multisheet CSV convention used by the
   sensitivity runner).

Tools:
- `tools/compute_aeff_values.py` — compute and cache A_eff lookup
   results (evaluate VLOOKUP/formula-derived columns) and write numeric
   A_eff into the `A_eff` sheet for downstream pipelines that lack
   Excel formula evaluation.

Testing / validation
- Local pytest: 42 passed, 4 warnings (curve-fit). The reorganized
   test suite is arranged by domain for easier maintenance and CI
   diagnostics.

Release actions recommended
- Create an annotated tag `v7.0.0` and attach a short changelog using
   the summary above. This will produce a traceable release marker for
   CI and package consumers.


Verification & notes:

- The GUI export saves workbooks in a background thread; automated
   headless tests may need to wait for the background save to complete or
   run the export synchronously for deterministic testing.
- To verify: open the GUI, apply a `Variable X% Y keV` preset to a few
   selected MMs and use Export → Save as new file; confirm `A_eff` column B
   contains numeric values (not the textual `gaussian(...)` expression) and
   the vignetting sheets' `C2` cell contains the parsed energy.

### Requirements
- Python 3.8 or higher
- pip (Python package manager)

### Setup

1. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

   Required packages:
   - numpy
   - pandas
   - matplotlib
   - openpyxl

2. **Verify installation**:
   ```bash
   python3 --version  # Should be 3.8+
   python3 -c "import numpy, pandas, matplotlib, openpyxl; print('All dependencies installed successfully')"
   ```

## Quick Start

### Using the GUI (Recommended)

1. **Launch the GUI**:
   ```bash
   python3 gui_distributions.py
   ```

   **macOS trackpad note:** If you use tap-to-click, the GUI now treats button activation more leniently (it won’t require a perfectly stationary press/release). For dropdowns (comboboxes), clicking anywhere on the field should open the list (not only the small arrow).


2. **Load your Excel file**:
   - Click "Load Excel File"
   - Navigate to the `Distributions/` folder
   - Select an Excel file with MM configuration

3. **Select Mirror Modules**:
   - Go to "MM Configuration" tab
   - Check/uncheck MMs to include in analysis
   - Use filters to select specific rows or MM numbers

4. **Generate PSF Data**:
   - Go to "PSF" → "Generate" tab
   - Choose distribution type: gaussian or pseudo-voigt
   - Set parameters for each MM characteristic
   - Click "Generate PSF Data"

5. **Export Results**:
   Short User Manual (GUI)
   ----------------------

   Follow these concise steps to use the GUI for typical generation and export tasks, including the latest v5 behaviors:

   1. Launch and load
      - Run `python3 gui_distributions.py` and click **Load Excel File** to open a workbook from `Distributions/`.

   2. Select MMs
      - In **MM Configuration** pick the mirror modules to modify (or `Select All`). Use filters to narrow selection.

   3. Choose A_eff preset or fixed value
      - In the **A_eff** tab select **Standard Distribution** to choose a preset from the workbook, or **Fixed Value** to enter a single numeric weight.
      - If the preset name contains an energy token (e.g. `1 keV`) or you choose a free-energy from the combobox, that numeric energy will be written into cell `C2` of the vignetting sheets on export.

   4. Apply & preview
      - Click **Apply to Selected MMs** to evaluate the preset for the selected MMs and preview the numeric values in the table. For percent-variable presets (e.g. `Variable 10% 1 keV`) the GUI synthesizes them to explicit gaussian forms so the preview shows sampled numeric per-MM values.

   5. Export with vignetting (optional)
      - In **A_eff**, check **Apply vignetting factors when exporting** to enable copying vignette columns into the export.
      - In **Preview / Export** choose **Save as new file**, then **Export to Excel**. The GUI export evaluates presets per-MM and writes numeric A_eff into column B of the `A_eff` sheet and clears column C. The vignetting sheets will have their `C2` cell set to the selected/parsed energy.

   6. Verify output
      - Open the saved workbook and confirm:
        - `A_eff` sheet: column B contains numeric per-MM A_eff values (not textual expressions).
        - Vignetting sheets: cell `C2` equals the parsed energy (if present).

   Notes and tips
    - The GUI save runs in a background thread; allow a short moment before opening the exported file or use the Preview to inspect results.
    - CLI runs via `main.py` keep the legacy behavior (adjusted A_eff written to column C). Use `main.py` for automated pipelines.
    - If deterministic, repeatable sampling is required for tests, use the same workbook filename and preset; the sampling is deterministic per-index and reproducible across runs given the same inputs.

   - Go to "Export" tab
   - Choose export mode (new file or update current)
   - Click "Export to Excel"
   - Files are saved in `Distributions/` folder
   - Context menu (right-click) on plots: Export PSF / Encircled Energy / FITS / EEF CSV / Fit Parameters CSV. EEF CSV files are written to `CustomPSFs/` as `E2E_EEF_YYYYMMDD_HHMMSS.csv`.

   ## Batch Processing & CLI (Headless)

   This project supports headless batch processing for automated generation of PSF products, packaging, and aggregated summaries. Use the CLI `main.py` for single-run, single-config, or packaged exports. The important flags and workflows are described below.

   Basic single-run examples

   - Coarse (fast) export and package:

   ```bash
   python3 main.py --file Distributions/YourWorkbook.xlsx --export-package --mode coarse
   ```

   - Fine (high-resolution) export and package — note: can be compute-heavy and may take significant time:

   ```bash
   python3 main.py --file Distributions/YourWorkbook.xlsx --export-package --mode fine
   ```

   - Run only a single mirror-module configuration index (useful for debugging):

   ```bash
   python3 main.py --file Distributions/YourWorkbook.xlsx --export-package --single-config 1
   ```

   Background/long-run tips

   - To run long fine-mode exports unattended, use `nohup` or a terminal multiplexer such as `screen` or `tmux`:

   ```bash
   nohup python3 main.py --file Distributions/YourWorkbook.xlsx --export-package --mode fine &
   # or
   screen -S e2e_run
   python3 main.py --file Distributions/YourWorkbook.xlsx --export-package --mode fine
   # detach with Ctrl-A D
   ```

   - If your runs are CPU-heavy, consider running on a machine with more CPU cores or limiting the grid resolution in `main.py` for fine-mode if you need faster turnaround.

   Export package contents and verification

   - Packages are written to `Exports/<TIMESTAMP>/` and include the packaged workbook (used as the authoritative source when aggregating A_eff sums), Figures in `Figures/`, and FITS files in `CustomPSFs/`.
   - The CLI prefers the workbook copied into the package when computing `Aeff_sum_orig` and `Aeff_sum_mod` so aggregated results match the packaged workbook.
   - After a successful package export, verify:
      - PNGs in `Figures/` (coarse = 320×320 px, fine = 2062×2062 px)
      - FITS in `CustomPSFs/` (grid dimensions equal to the requested pixel size)
      - Aggregated Excel `E2E_EEF_and_fitparams_*.xlsx` present in the package.

   Repairing existing packages

   If you have existing `Exports/` packages created before these fixes, use the included repair/diagnostic scripts at the repository root to inspect and patch packages:

   - `.inspect_export.py` — inspect a package and list mismatches between packaged workbook and recorded aggregated sums.
   - `.diagnose_aeff.py` — diagnostics to identify formula-only A_eff columns and missing cached values.
   - `.patch_fitparams.py` — patch per-config `fitparams_aeffloss.xlsx` files inside packages when sums differ.
   - `.patch_aggregated.py` — repair aggregated workbook rows that used the wrong source for A_eff sums.

   Example: run the inspect script and, if needed, apply a patch (these are best run interactively to confirm changes):

   ```bash
   python3 .inspect_export.py Exports/20260416_124558
   python3 .patch_fitparams.py Exports/20260416_124558 --dry-run
   python3 .patch_fitparams.py Exports/20260416_124558
   ```

   Notes on formula evaluation

   - Many input workbooks use Excel formulas (VLOOKUP/XLOOKUP or textual preset expressions). When Excel cached numeric values are missing, a best-effort internal evaluator attempts to resolve common patterns and populate numeric `A_eff` values during export.
   - The helper script `tools/compute_aeff_values.py` can be used to precompute and write numeric `A_eff` columns into workbooks so downstream runs do not rely on formula evaluation logic.

   Troubleshooting

   - If an exported package's aggregated `Aeff_sum_mod` appears incorrect, re-run the inspect script and check whether the packaged workbook's `A_eff` sheet contains formulas without cached values. If so, use `tools/compute_aeff_values.py` or re-export from Excel to populate cached values.
   - Fine-mode runs can be very long; if interrupted you can re-run the same command and the package writer will either pick up partial results or re-generate the package depending on the stage at interruption.


### Using Command Line

```bash
python3 main.py -f Distributions/your_file.xlsx
```

If you don’t pass `--output`, a window opens with interactive export shortcuts.
If you pass `--output`, the combined figure is saved to that path and the script exits without opening a window.

## Documentation updates (2026-02-03)

- Added module-level documentation to core modules and a contributor guide.
- Introduced `DOCS_SUMMARY.md` and `CONTRIBUTING.md` to the repository root.
- Cleaned up transient debug scripts from `scripts/` and added placeholders where helpful.

## Release v4 (2026-02-03)

- Fixed dz→dm projection and ensured per‑MM dz outputs are correct when
   running `--log-dz` (proper polar/cartesian projection and use of
   alignment/gravity/thermal rotations).
- Re-ordered vignetting application to run after A_eff initialization and
   added explicit bookkeeping: `aeff_base`, `aeff_adjusted`, and
   `aeff_vig_factor` so plots and aggregations use the adjusted effective
   area.
- Improved vignetting parsing to support both two-column (delta→factor)
   sheets and per-position columns; MM300 now receives the combined
   vignette factor when multiple factors apply (e.g. 0.1×0.1→0.01).
- Updated rotation totals computation to include: alignment direct polar +
   projected contributions from gravity/thermal rotx/roty + direct
   gravity/thermal polar terms. Removed deprecated `d_align_rotx` and
   `d_align_roty` usage.
- GUI changes: added `Apply vignetting factors when exporting` checkbox
   in the `A_eff` tab; when enabled, the selected preset column from the
   vignetting sheets is copied into column B during export (matches by
   explicit column letter or header substring).
- Extended the interactive plot context menu to export fit parameters CSV,
   and overlaid the aggregated pseudo-Voigt fit on the right-hand EEF plot.
- Documentation: substantial docstring pass across `main.py`,
   `gui_distributions.py`, `distributions_rotated.py`, and
   `optimize_mm_rows.py`; added `DOCS_GUI.md` with examples and generated
   thumbnail screenshots in `Figures/`.
- Tests: updated tests and workbook references; full local test suite
   passes (34 tests) after these changes.

See `DOCS_GUI.md` for usage examples, screenshots and troubleshooting notes.

See `DOCS_SUMMARY.md` for a short map of key files and recommended next steps.

## Release v5 (2026-02-07)

- GUI A_eff export semantics changed: when exporting from the GUI, a
   selected *standard* A_eff preset is now evaluated per-MM and the resulting
   numeric A_eff values are written into column B of the `A_eff` sheet. The
   GUI export path explicitly clears column C so adjusted/vignetted A_eff
   values are not written by the interactive export. The CLI/main program
   (`main.py`) preserves the previous behavior and continues to populate
   column C with adjusted A_eff when run non-interactively.
- Percent-variable presets are synthesized into explicit gaussian forms when
   loaded from the workbook (e.g. a preset named "Variable 10% 1 keV" with
   a Values cell of `L` becomes internally `gaussian(L,10%*L)`) so the
   evaluator can sample correctly.
- Preset-energy parsing improved: when a preset name contains an energy
   token such as `1 keV`, the GUI now parses that numeric energy and writes
   it into cell `C2` of the vignetting sheets at export time. The export
   routine prefers an explicit free-energy selection (combobox) but falls
   back to the `keV` token in the preset name if present.
- The preset evaluator (`_evaluate_aeff_preset_for_mm`) is used at export
   time so per-MM draws are consistent with the Apply action; a numeric
   fallback uses the `A_eff_base` value and the parsed percent when the
   direct evaluation fails.
- Removed noisy debug prints and replaced them with Python `logging` calls
   across the GUI module so runtime logs are quieter in normal operation.
- Misc: small repository cleanup (removed stale preview workbooks under
   `Distributions/` and created a zipped backup). The branch `cleanup/v5`
   containing these updates was merged into `main` on 2026-02-07.

Verification & notes:

- The GUI export saves workbooks in a background thread; automated
   headless tests may need to wait for the background save to complete or
   run the export synchronously for deterministic testing.
- To verify: open the GUI, apply a `Variable X% Y keV` preset to a few
   selected MMs and use Export → Save as new file; confirm `A_eff` column B
   contains numeric values (not the textual `gaussian(...)` expression) and
   the vignetting sheets' `C2` cell contains the parsed energy.

## Docs Index

For detailed documentation, see:

| Topic | Document |
|-------|----------|
| Core CLI & analysis | [README.md](README.md) |
| GUI usage & screenshots | [DOCS_GUI.md](DOCS_GUI.md) |
| API reference | [DOCS_SUMMARY.md](DOCS_SUMMARY.md) |
| Contributing guidelines | [CONTRIBUTING.md](CONTRIBUTING.md) |
| Release history | [RELEASE_NOTES.md](RELEASE_NOTES.md) |
| Quick start guide | [QUICKSTART.txt](QUICKSTART.txt) |
| Sensitivity pipeline | [SENSITIVITY_QUICKSTART.txt](SENSITIVITY_QUICKSTART.txt) |

### Core Modules
- **Core loader & CLI:** [main.py](main.py)
- **GUI application:** [gui_distributions.py](gui_distributions.py)
- **Distribution utilities:** [distributions_rotated.py](distributions_rotated.py)
- **Row optimizer & helpers:** [optimize_mm_rows.py](optimize_mm_rows.py)

### Module reference

- `main.py`: Command-line entrypoints and Excel I/O helpers. Implements
   robust readers for `MM_PSF` and `A_eff` tables, spreadsheet-preserving
   export paths, PSF parameter conversion (arcsec→meters), vignetting
   application, and the headless plot/export options used in CI and
   batch sensitivity runs. Key helper groups: workbook parsing,
   distribution sampling, vignetting evaluation, MM configuration join
   logic (MM#↔Position), and figure export utilities.

- `gui_distributions.py`: Tkinter GUI for interactive generation and
   exporting of per-MM distributions. Loads standard presets from the
   workbook, allows per-data-type distribution editing (A_eff, MM_PSF,
   Alignment, Thermal, Gravity), previews sampled tables, and performs
   export with optional vignetting copy. Use `tools/compute_aeff_values.py`
   to prepare numeric A_eff workbooks for headless pipelines.

## Directory Structure

```
NewAthenaE2EPSF_v3/
├── README.md                      # This file
├── requirements.txt               # Python dependencies
├── main.py                        # CLI analyzer and plotter
├── gui_distributions.py           # Interactive GUI application
├── distributions_rotated.py       # Core distribution functions
├── optimize_mm_rows.py            # MM optimizer + placement + Excel writer
├── Distributions/                 # Excel spreadsheet location
│   ├── TestDistribution.xlsx    # Example file
│   └── your_data.xlsx             # Your files go here
├── Figures/                       # Exported plots location
│   ├── E2E_PSF_*.png
│   └── Encircled_Energy_*.png
└── scripts/                       # Helper scripts for export/diagnostics
```

## Usage

### GUI Application

The GUI provides a complete workflow for generating and managing MM PSF data:

#### 1. MM Configuration Tab
- **Load Excel**: Import existing MM configuration
- **Selection**: Choose which MMs to modify
- **Filters**: Filter by Row # or MM #
- **Sort**: Click column headers to sort

**macOS trackpad tip:** In the MM Configuration table you can toggle an MM by clicking anywhere in its row. You can also use keyboard selection + `Space` to toggle the currently highlighted rows.

#### 2. PSF → Select Tab
Information about using the Generate tab for PSF data.

#### 3. PSF → Generate Tab

**Distribution Type Selection:**
```
┌─────────────────────────────────────────────────────────────┐
│  Distribution Type: [ gaussian ▼ ]  or  [ pseudo-voigt ▼ ] │
└─────────────────────────────────────────────────────────────┘
```

**For Gaussian Distribution:**
- eta parameters are hidden
- Only PSF parameters (m_rad, m_azi, sigma_rad, sigma_azi) are shown

**For Pseudo-Voigt Distribution:**
- eta parameter row appears
- Configure eta distribution (fixed/gaussian/uniform)
- eta values are automatically clamped to [0.0, 1.0]

**Parameter Configuration:**
Each parameter has three distribution options:

1. **Fixed** (default):
   - Single value applied to all MMs
   - Second input field is hidden

2. **Gaussian**:
   - Mean: Center value of the distribution
   - Sigma: Standard deviation (spread)
   - Each MM gets a random value from N(mean, sigma²)

3. **Uniform**:
   - Min: Lower bound
   - Max: Upper bound
   - Each MM gets a random value uniformly distributed in [min, max]

#### 4. Other Data Types
- **Alignment Errors**: Generate MM alignment perturbations
- **Gravity Offload**: Simulate gravity effects
- **Thermal Deformation**: Model thermal expansion/contraction

#### 5. Export Tab
- **Current File**: Update the loaded Excel file
- **New File**: Save to a new file in `Distributions/` folder
- **Preview**: See what will be exported

### Command Line Usage

```bash
python3 main.py [OPTIONS]
```

**Options:**
- `-f, --file FILE`: Path to Excel file (default: `Distributions/TestDistribution.xlsx`)
- `--normalize`: Normalize PSF to unit integral
- `--no-normalize`: Disable normalization
- `--output FILE`: Save combined plot to file (in `Figures/` folder)
- `--mode {coarse,fine}`: Runtime mode. Controls plotting + optimization speed/accuracy.

Additional export and metadata options:

- `--author NAME` (exporter script): write `AUTHOR` into FITS header when exporting via `scripts/export_e2e_fits.py`.
- `--contact EMAIL` (exporter script): write `CONTACT` into FITS header (falls back to git `user.email` or ivo.ferreira@esa.int).
- `--orcid ORCID` (exporter script): write `ORCID` into FITS header.

GUI notes:

- Right-click on the interactive plots to open a context menu. Options include:
   - Export PSF Plot (PNG)
   - Export Encircled Energy Plot (PNG)
   - Export FITS (Primary HDU-only, big-endian IEEE64) — header includes `TOT_AEFF`, `INTG_Z`, `PIXAS*`, `PIXM*`, `CDELT*`, `AUTHOR`, `CONTACT`, `ORCID`, `INPUTFN`.
   - Export EEF CSV (writes `CustomPSFs/E2E_EEF_YYYYMMDD_HHMMSS.csv` with percentage/diameter columns for best/origin/optimized curves when present).
   - Export Fit Parameters CSV (writes `CustomPSFs/E2E_fit_params_YYYYMMDD_HHMMSS.csv` with aggregated pseudo-Voigt fit parameters).

Keyboard shortcuts in the interactive window:

- `p` or `1`: Export PSF plot to PNG
- `e` or `2`: Export Encircled Energy plot to PNG
- `f` or `3`: Export aggregated E2E PSF to FITS
- `c` or `4`: Export Encircled Energy Function data to CSV (`CustomPSFs/`)
- `s` or `5`: Export fit parameters to CSV (`CustomPSFs/`)
- `--optimize`: Enable row-wise MM# assignment optimization to minimize HEW (uses `--mode`).
   - Note: `--optimise` is accepted as an alias.
- `--placement [{cross,x_axis,elliptical}]`: Apply a placement strategy.
   - Runs placement-only when used alone (writes `*_placed.xlsx` and overlays it).
   - Seeds the optimizer when used together with `--optimize`.
   - If `--optimize` is used without `--placement`, the default seed is `elliptical`.

**Examples:**

1. **Basic analysis**:
   ```bash
   python3 main.py -f Distributions/my_data.xlsx
   ```

2. **Normalized PSF**:
   ```bash
   python3 main.py -f Distributions/my_data.xlsx --normalize
   ```

3. **Coarse mode** (fast, for quick previews):
   ```bash
   python3 main.py -f Distributions/my_data.xlsx --mode coarse
   ```

4. **Fine mode** (higher sampling density):
   ```bash
   python3 main.py -f Distributions/my_data.xlsx --mode fine
   ```

6. **Save output figure**:
   ```bash
   python3 main.py -f Distributions/my_data.xlsx --output Figures/my_analysis.png
   ```

7. **Placement only** (writes `*_placed.xlsx` and overlays it):
   ```bash
   python3 main.py -f Distributions/my_data.xlsx --placement cross
   python3 main.py -f Distributions/my_data.xlsx --placement x_axis
   python3 main.py -f Distributions/my_data.xlsx --placement elliptical
   ```

8. **Optimize MM assignments**:
   ```bash
   python3 main.py -f Distributions/my_data.xlsx --optimize
   ```
   This finds the best assignment of MM# to physical locations within each row to minimize HEW. The optimized configuration is saved to `*_optimised.xlsx`.

9. **Optimize with fine mode**:
   ```bash
   python3 main.py -f Distributions/my_data.xlsx --mode fine --optimize
   ```

10. *(Note: `extra-fine` mode has been removed; use `--mode fine` for high-accuracy runs.)*

11. **Optimize seeded by a placement**:
   ```bash
   python3 main.py -f Distributions/my_data.xlsx --mode fine --optimize --placement elliptical
   ```

**Interactive Controls:**
When the plot window opens:
- `p` or `1`: Export PSF plot to PNG
- `e` or `2`: Export Encircled Energy plot to PNG
- `Right-click` (or `Ctrl+click` on Mac): Show context menu
- `h`: Show help

Figures are automatically timestamped and saved to `Figures/` folder.

## Plot legend notes

The encircled-energy legend displays `HEW_x` and `HEW_y` in arcsec.
These are computed on the full **summed PSF** (after aggregating all distributions) by:
- Computing the 1D marginals $P_x(x)=\int Z(x,y)\,dy$ and $P_y(y)=\int Z(x,y)\,dx$
- Finding the **smallest interval anywhere** containing 50% of the marginal energy


## MM Position Optimizer

### Overview
The row-wise MM optimizer finds the best assignment of MM numbers to physical locations within each row to minimize the system's Half-Energy Width (HEW). Unlike traditional optimization that moves physical positions, this optimizer keeps locations (x, y, z, r) fixed and swaps which MM# is assigned to each position.

### How It Works
1. **Row-Based Optimization**: Each row is optimized independently, preserving row structure
2. **MM# Swapping**: Within each row, MM# assignments are permuted while locations remain fixed
3. **Parameter Lookup**: For each assignment, PSF parameters are retrieved from the MM_PSF sheet based on MM#
4. **HEW Evaluation**: Fast approximate HEW computation (~60×40 polar grid) tests each permutation
5. **Iterative Improvement**: Random swaps are accepted only if they reduce HEW
6. **Output**: Optimized MM configuration saved to `*_optimised.xlsx` with all sheets preserved

### Placement Strategies (Seeding / Placement-Only)

The `--placement` option provides deterministic “first guess” configurations. You can use it:
- **Alone** to write a `*_placed.xlsx` file and overlay it on the plots.
- **With `--optimize`** to seed the optimizer (i.e., start from a structured configuration instead of the input order).

All placement strategies follow the same pattern:
1. Compute a per-MM “badness” score within each row (lower = better).
2. Sort MMs in that row by score.
3. Assign the best/worst MMs to different slot-angle patterns depending on the strategy.

**How “best” vs “worst” is ranked**

The per-MM score used for placement/seed is **shift-aware**:
- It uses the intrinsic PSF terms from the MM’s own parameters (from `MM_PSF`).
- It also includes the expected centroid shifts due to perturbations (Alignment / Gravity offload / Thermal).
- Those perturbations are applied **per position (slot)**, not per MM: when MMs are swapped, the deltas stay with the slot.

This means a module can rank “worse” if a given slot’s deltas would push its spot farther from the nominal center.

**Strategies**

- `cross`:
   - Places the best MMs on a 90° “cross-like” pattern (closer to the ±x / ±y directions).
   - Intuition: spread the best contributors across orthogonal directions to reduce anisotropy and provide a robust seed.

- `x_axis`:
   - Places the best MMs preferentially near the ±x direction.
   - Within that, it alternates slots above/below the x-axis to avoid clustering everything on one side.

- `elliptical`:
   - Within each row, assigns the best MMs to slots closer to the x-axis and the worst MMs to slots closer to the y-axis.
   - Intuition: creates an “ellipse-like” ordering where the row’s strongest performers concentrate near the horizontal axis.

### Usage

**Command Line**:
```bash
python3 main.py -f Distributions/input.xlsx --mode coarse --optimize
```

**Options**:
 - `--mode {coarse,fine}`: Controls optimization budget and sampling density
- `--optimize`: Enable row-wise MM# assignment optimization

**Output Files**:
- Input: `Distributions/my_data.xlsx`
- Output: `Distributions/my_data_optimised.xlsx`
- Plots: `Figures/E2E_PSF_*_optimised_*.png`

**Note:** Files like `*_optimised.xlsx` and `*_placed.xlsx` are generated outputs and are not shipped with this project. The `Distributions/` folder in the repo only includes `TestDistribution.xlsx` by default.

**Performance**:
- Typical runtime: 15-30 seconds for ~150 MMs
- Speed optimizations: Ultra-coarse polar grid (60×40), skip iterative focus refinement during permutation testing
- Memory: Minimal, processes one row at a time

### When to Use
- **Varying MM PSF parameters**: When different MM# have different PSF characteristics in MM_PSF sheet
- **Row constraints**: When MMs within a row must stay in that row but can swap positions
- **HEW minimization**: When you want to find the best spatial arrangement without redesigning the system

### Example Workflow
```bash
# 1. Generate MM configuration with variable PSF parameters
python3 gui_distributions.py
# (Use GUI to create varied MM_PSF parameters)

# 2. Analyze baseline HEW
python3 main.py -f Distributions/my_config.xlsx
# HEW: 15.2 arcsec (baseline)

# 3. Optimize MM# assignments
python3 main.py -f Distributions/my_config.xlsx --optimize
# HEW: 14.1 arcsec (optimized, 7.2% improvement)

# 4. Compare results
python3 main.py -f Distributions/my_config.xlsx
python3 main.py -f Distributions/my_config_optimised.xlsx
```

## Standard MM_PSF Distributions

### Overview
The GUI supports loading predefined MM_PSF distributions from an Excel table inside your workbook. This is the **source of truth** for what appears in the “Standard” preset dropdown.

In **Standard mode**, the GUI:
- Reads the preset table from the currently loaded Excel file
- Populates the `sigma_rad` / `sigma_azi` (and pseudo-voigt `alpha_*`) controls from that table
- Forces `m_rad` and `m_azi` to **fixed 0** for consistency (the standard table does not define them)

### Preset Modes

**1. Standard Mode** (default):
- Loads predefined distributions from the **MM_PSF sheet table starting at cell K1**.
- The dropdown is populated directly from the “Preset Name” column.
- The sigma/alpha fields are filled from that row.
- Use this when you want repeatable, centrally managed presets.

**2. Free Mode**:
- Full manual control over all PSF parameters
- Choose distribution type (gaussian/pseudo-voigt)
- Set each parameter's distribution (fixed/gaussian/uniform)
- Configure means, sigmas, min/max values

### Using Standard Presets

**GUI Workflow**:
1. Launch GUI: `python3 gui_distributions.py`
2. Load Excel file with MM configuration
3. Go to PSF → Select tab
4. Choose "Standard" mode (default)
5. Click dropdown next to "Load from Excel (K1):"
6. Select preset (e.g., `Fixed Asym Gaussian 8" (ratio 7:1)`)
7. The Generate tab fields update automatically
8. Click "Generate PSF Data" to apply to all selected MMs
9. Click "Generate PSF Data" to apply to all selected MMs

**Preset naming**:
- The name is primarily for humans (it’s what appears in the dropdown).
- The GUI does **not** “interpret” the name to compute asymmetric sigma values.
- If a table cell is blank and the name contains something like `10% Variable ... 8"`, the GUI may derive a symmetric Gaussian fallback — but for production usage you should put explicit values/specs in the table.

**Parameter Derivation**:
- **HEW → Sigma (symmetric Gaussian)**: For a circular (symmetric) 2D Gaussian, the **HEW diameter** (50% encircled-energy diameter) maps to
   σ = HEW / (2√(2ln2)) ≈ HEW / 2.355.
- **Asymmetric Gaussian note**: For an elliptical (asymmetric) 2D Gaussian, the 50% encircled-energy diameter depends on **both** axis sigmas (and the axis ratio), so there is no single closed-form σ = HEW/2.355 mapping. For asymmetric Gaussian presets (e.g. “Asym Gaussian 8\" (ratio 7:1)”), the standard preset table should therefore contain sigma values that already match the desired **encircled-energy HEW diameter**.
- **Percent Variability**: "10%" means sigma_sigma = mean_sigma × 0.10
- **Alpha Fallback**: If alpha not in preset name, uses matching fixed preset or default 0.77/0.29
- **Auto-Clamping**: All alpha values automatically clamped to [0, 1] range

**Recommended sigma values (arcsec) for common presets**

These values are consistent with the plotted green “HEW … diameter” (50% encircled-energy diameter) for a single Gaussian PSF:

- Symmetric Gaussian, 4.3" HEW: `sigma_rad = sigma_azi = 1.826058874`
- Symmetric Gaussian, 8" HEW: `sigma_rad = sigma_azi = 3.397323952`

For asymmetric Gaussians, the values depend on the axis ratio and are computed so that the 50% **encircled-energy** diameter equals the quoted HEW:

- Asymmetric Gaussian, 4.3" HEW, ratio 4:1:
   - `sigma_major = 2.963740901247624`
   - `sigma_minor = 0.740935225311906`
- Asymmetric Gaussian, 8" HEW, ratio 7:1:
   - `sigma_major = 5.797205089162605`
   - `sigma_minor = 0.8281721555946578`

If your preset uses `sigma_rad` for the “major” axis and `sigma_azi` for the “minor” axis, put the values in that order; otherwise swap them.

### Adding Custom Presets

**Excel Table Structure** (MM_PSF sheet, starting at cell K1):

Columns (left → right):
1. Preset Name (text)
2. `sigma_rad` spec
3. `sigma_azi` spec
4. `alpha_rad` spec (pseudo-voigt only; otherwise `-` or blank)
5. `alpha_azi` spec (pseudo-voigt only; otherwise `-` or blank)

Each “spec” cell can be one of:
- A plain number (interpreted as **fixed**) e.g. `3.397323952`
- A Gaussian spec: `gaussian(mean, sigma)` where `mean` and `sigma` may include simple arithmetic and `%`
   - Example: `gaussian(3.397, 10%*3.397)`
- A Uniform spec: `uniform(min, max)`

Notes:
- Values are in **arcsec**.
- In Standard mode the GUI forces `m_rad` and `m_azi` to fixed 0 (table does not define them).

**Common pitfalls (read this if a preset doesn’t load as expected)**

- **You must reload the workbook in the GUI**: the preset dropdown is populated from the Excel file you load. If you edit the table in Excel, save it, then re-load that workbook in the GUI.
- **Don’t leave “spec” cells blank**: blank `sigma_rad`/`sigma_azi` cells can trigger fallback behavior (or result in missing UI updates). Prefer explicit numeric/spec values.
- **Use dot decimals**: write `3.397`, not `3,397`.
- **Use the supported syntax only**:
   - Fixed: `3.397323952`
   - Gaussian: `gaussian(3.397, 10%*3.397)`
   - Uniform: `uniform(3.0, 4.0)`
   Anything else (extra text, units inside the cell, etc.) may be ignored.
- **Keep units consistent**: the table values are **arcsec**; don’t enter microns or radians.
- **Alpha columns for Gaussians**: for Gaussian presets, use `-` (dash) or leave alpha cells empty. Alpha values are only applied for pseudo-voigt.
- **Asymmetric Gaussian “HEW” is not σ=HEW/2.355**: if the name says “Asym Gaussian 8\" (ratio 7:1)”, you still must provide sigma values that match the 50% encircled-energy diameter (see the reference values above).
- **If the preset is pseudo-voigt but alpha doesn’t appear**: check that the preset name includes “pseudo-voigt” (or “voigt”) so it’s classified correctly.

1. Open your Excel file in `Distributions/`
2. Locate the standard distribution table (typically near cell K1)
3. Add/edit rows with the desired numeric/spec values
4. Save Excel file
5. In the GUI, load that Excel file (the preset dropdown is refreshed on load)

### Benefits
- **Consistency**: All MMs use the same baseline preset
- **Speed**: One-click loading vs manual parameter entry
- **Accuracy**: For symmetric Gaussians, the HEW→σ mapping is direct; for asymmetric Gaussians, you can author the exact sigma values that match the plotted encircled-energy HEW.
- **Flexibility**: Mix standard presets with custom free-mode configurations

## Distribution Types

### Gaussian Distribution
Standard 2D Gaussian function:

```
G(x,y) = A × exp(-0.5 × r²)
```

where r² is the rotated squared distance from the center.

**Characteristics:**
- Sharp central peak
- Exponential decay in tails
- Suitable for well-aligned, stable systems

### Pseudo-Voigt Distribution
Product of two independent 1D Pseudo-Voigt functions (azimuthal and radial):

```
PV(azi, rad) = PV_azi(azi) × PV_rad(rad)
```

Each 1D component is a weighted sum of normalized Gaussian and Lorentzian:
```
PV_1D(x) = α × L(x) + (1-α) × G(x)
```

where:
- `α` (alpha) is the mixing parameter [0, 1] (note: different from traditional eta notation)
- `α = 0`: Pure Gaussian (sharp peak)
- `α = 0.5`: Balanced mix  
- `α = 1`: Pure Lorentzian (heavy tails)
- Independent alpha_azi and alpha_rad for asymmetric PSFs

**Alpha Values Guide:**
```
α = 0.0  ████████████░░░░░░░░  Pure Gaussian (sharp core)
α = 0.2  ████████████▓░░░░░░░  Mostly Gaussian
α = 0.5  ████████████▓▓▓░░░░░  Balanced (50/50 mix)
α = 0.7  ████████████▓▓▓▓▓░░░  More tails (scattering)
α = 1.0  ████████████▓▓▓▓▓▓▓▓  Pure Lorentzian (heavy tails)
         └───Core───┘└─Tails─┘
```

**Use Cases:**
- **Low alpha (0.0 - 0.3)**: Good alignment, minimal scattering  
- **Medium alpha (0.3 - 0.7)**: Moderate scattering effects
- **High alpha (0.7 - 1.0)**: Significant scattering, degraded optics
- **Asymmetric (alpha_rad ≠ alpha_azi)**: Directional scattering or optical aberrations

### Aggregated Radial Fit: Modified Pseudo-Voigt

The current end-to-end fit in `main.py` uses a modified pseudo-Voigt model applied to the azimuthal-average radial intensity profile of the aggregated PSF.
The model is not a simple Gaussian; it combines a narrow Gaussian core with a broader wing-shaped term and normalizes the mixture so the amplitude `A` remains the overall intensity scale.

The mathematical form is:

```
G(r; Γ_c) = exp(-4 ln 2 (r / Γ_c)^2)
C(r; Γ_w) = [1 + a (2 r / Γ_w)^2]^{-β}
where a = 2^(1/β) - 1
mix = (1 - η) G(r; Γ_c) + η × scalar × C(r; Γ_w)
norm = (1 - η) + η × scalar
I(r) = A × mix / norm
```

Parameters:

- `A`: overall intensity scale of the radial profile. This is the fitted amplitude of the mean radial intensity.
- `Γ_c` (`Gamma_core`): core width in arcseconds. Controls the narrow Gaussian-like peak width.
- `Γ_w` (`Gamma_wing`): wing width in arcseconds. Controls the broader wing/tail scale.
- `η` (`eta`): mixing fraction between core and wing. `η = 0` gives pure Gaussian core, `η = 1` gives a wing-dominated shape.
- `β` (`beta`): wing shape exponent. `β = 1` produces a Lorentzian-like tail, larger values produce faster-decaying wing tails.
- `scalar`: additional scaling factor applied to the wing component before normalization. Values above 1 amplify the wing relative to the Gaussian core, values below 1 reduce the wing amplitude.

Because the model normalizes by `(1 - η) + η scalar`, the fit preserves the overall amplitude while allowing the wing component to have its own relative strength.

This aggregated fit is performed on the azimuthally-averaged radial intensity derived from the PSF, and the result is saved as `Figures/E2E_fit.png`.

## GUI User Manual

### Overview

The GUI provides an intuitive interface for generating PSF configurations with support for standard presets, manual parameter control, and multi-data type generation. The interface is organized into tabs for different workflows.

**Main Window Layout:**
```
┌─────────────────────────────────────────────────────────────────┐
│  New Athena E2E - PSF Configuration Tool              [_][□][×] │
├─────────────────────────────────────────────────────────────────┤
│  [Load Excel File]  Current: None                               │
├─────────────────────────────────────────────────────────────────┤
│  ┌────────────────────────────────────────────────────────────┐ │
│  │ [MM Configuration] [PSF] [Other Data Types] [Export]       │ │
│  │                                                             │ │
│  │                  (Tab content area)                         │ │
│  │                                                             │ │
│  └────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────┘
```

### Step-by-Step Tutorial

#### Step 1: Launch and Load File

1. **Start the GUI:**
   ```bash
   cd "/Users/.../NewAthenaE2EPSF_v2"
   python3 gui_distributions.py
   ```

2. **Load Excel file:**
   - Click **"Load Excel File"** button at top
   - Navigate to `Distributions/` folder
   - Select an existing file (e.g., `TestDistribution.xlsx`)
   - File path displays in header: "Current: Distributions/your_file.xlsx"
   - MM Configuration tab automatically populates with data

**Screenshot location: Initial window and file selection dialog**

#### Step 2: MM Configuration Tab

The first tab shows all MMs from your Excel file in a spreadsheet view.

```
┌─────────────────────────────────────────────────────────────────┐
│  MM Configuration Tab                                           │
├─────────────────────────────────────────────────────────────────┤
│  ☑ Select/Deselect All                                          │
│                                                                  │
│  Filter Options:                                                │
│    Row #: [All ▼]    MM #: [All ▼]                              │
│                                                                  │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │ ☑ | Petal# | Row# | MM# | x_MM[m] | y_MM[m] | z_MM[m] |...│ │
│  │ ☑ |   1    |  1   |  1  | 0.1234  | 0.5678  | 7.8900  |...│ │
│  │ ☑ |   1    |  1   |  2  | 0.1235  | 0.5679  | 7.8901  |...│ │
│  │ ☑ |   1    |  2   |  3  | 0.2345  | 0.6789  | 7.8902  |...│ │
│  │ ... (scrollable table with all MMs)                         │ │
│  └────────────────────────────────────────────────────────────┘ │
│                                                                  │
│  Selected: 145/150 MMs             [Apply Selection]            │
└─────────────────────────────────────────────────────────────────┘
```

**Features:**
- **Checkboxes**: Select which MMs to include in PSF generation
- **Select/Deselect All**: Toggle all checkboxes at once
- **Filters**: Show only specific rows or MM numbers
- **Sortable columns**: Click column headers to sort
- **Apply Selection**: Must click to confirm selection changes

**Common Actions:**
- Select all MMs in Row 1: Filter by Row# = 1, check all
- Exclude MM# 5-10: Filter by MM# range, uncheck, reset filter
- Sort by position: Click "x_MM [m]" header

**Screenshot location: MM Configuration tab with filters applied**

#### Step 3A: PSF Generation - Standard Mode (Recommended)

Use standard presets for quick, consistent PSF generation.

**Sub-tabs:** `[Select]` `[Generate]`

##### Select Sub-tab

```
┌─────────────────────────────────────────────────────────────────┐
│  PSF → Select                                                   │
├─────────────────────────────────────────────────────────────────┤
│  Preset Mode:                                                   │
│    ◉ Standard    ○ Free                                         │
│                                                                  │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │  Load from Excel (K1):                                      │ │
│  │                                                              │ │
│  │  [gaussian 15 arcsec HEW 10%          ▼]                   │ │
│  │                                                              │ │
│  │  Available presets:                                         │ │
│  │    • gaussian 10 arcsec HEW (fixed)                         │ │
│  │    • gaussian 15 arcsec HEW 10% (variable)                  │ │
│  │    • pseudo-voigt 12 arcsec HEW 0.77/0.29                   │ │
│  │                                                              │ │
│  │            [Load Standard Distribution]                      │ │
│  └────────────────────────────────────────────────────────────┘ │
│                                                                  │
│  Status: Ready to load preset                                   │
└─────────────────────────────────────────────────────────────────┘
```

**Steps:**
1. Ensure **"Standard"** radio button is selected (default)
2. Click dropdown to see available presets
3. Select desired preset (e.g., "gaussian 15 arcsec HEW 10%")
4. Click **"Load Standard Distribution"** button
5. Status updates: "Loaded preset: gaussian 15 arcsec HEW 10%"
6. Switch to **Generate** sub-tab to see auto-filled parameters

**Preset Types:**
- **Fixed presets** (no %): All MMs get identical values
- **Variable presets** (with %): Parameters vary per MM using Gaussian distribution
- **Pseudo-Voigt presets** (with ratio): Include alpha_rad/alpha_azi values

##### Generate Sub-tab (Standard Mode)

After loading a preset, the Generate tab shows the configured parameters:

```
┌─────────────────────────────────────────────────────────────────┐
│  PSF → Generate                                                 │
├─────────────────────────────────────────────────────────────────┤
│  Distribution Type: [gaussian  ▼]                               │
│                                                                  │
│  Parameters (Auto-filled from preset):                          │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │  m_rad [arcsec]:       [gaussian ▼]  [0.000]  [0.000]      │ │
│  │  m_azi [arcsec]:       [gaussian ▼]  [0.000]  [0.100]      │ │
│  │  sigma_rad [arcsec]:   [gaussian ▼]  [6.363]  [0.636]      │ │
│  │  sigma_azi [arcsec]:   [gaussian ▼]  [6.363]  [0.636]      │ │
│  │                                                              │ │
│  │  Note: Values auto-calculated from "15 arcsec HEW 10%"      │ │
│  │        σ_mean = 15/(2√(2ln2)) ≈ 6.363 arcsec                │ │
│  │        σ_sigma = 6.363 × 0.10 = 0.636 arcsec                │ │
│  └────────────────────────────────────────────────────────────┘ │
│                                                                  │
│                    [Generate PSF Data]                           │
│                                                                  │
│  Status: Parameters loaded from standard preset                 │
└─────────────────────────────────────────────────────────────────┘
```

**Key Points:**
- Parameters are **read-only** in standard mode (grayed out in actual GUI)
- Distribution types already set (fixed/gaussian/uniform)
- HEW-to-sigma conversion done automatically
- Click **"Generate PSF Data"** to apply to selected MMs

**Screenshot location: Generate tab with auto-filled standard preset parameters**

#### Step 3B: PSF Generation - Free Mode (Advanced)

For full manual control over all parameters.

##### Select Sub-tab (Free Mode)

```
┌─────────────────────────────────────────────────────────────────┐
│  PSF → Select                                                   │
├─────────────────────────────────────────────────────────────────┤
│  Preset Mode:                                                   │
│    ○ Standard    ◉ Free                                         │
│                                                                  │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │  Free Mode: Full manual control                             │ │
│  │                                                              │ │
│  │  Use the Generate tab to configure all parameters manually. │ │
│  │                                                              │ │
│  │  You can:                                                    │ │
│  │    • Choose distribution type (gaussian/pseudo-voigt)        │ │
│  │    • Set each parameter's distribution (fixed/gauss/uniform) │ │
│  │    • Configure means, sigmas, min/max values                 │ │
│  │    • Control alpha parameters for pseudo-voigt               │ │
│  └────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────┘
```

##### Generate Sub-tab (Free Mode)

```
┌─────────────────────────────────────────────────────────────────┐
│  PSF → Generate                                                 │
├─────────────────────────────────────────────────────────────────┤
│  Distribution Type: [pseudo-voigt ▼]                            │
│                                                                  │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │ Parameter       Distribution    Value 1      Value 2        │ │
│  ├────────────────────────────────────────────────────────────┤ │
│  │ m_rad [arcsec]  [fixed    ▼]   [0.000]      (hidden)        │ │
│  │ m_azi [arcsec]  [gaussian ▼]   [5.000]      [0.500]         │ │
│  │ sigma_rad ["]   [gaussian ▼]   [8.000]      [0.800]         │ │
│  │ sigma_azi ["]   [gaussian ▼]   [8.000]      [0.800]         │ │
│  │ alpha_rad       [gaussian ▼]   [0.770]      [0.050]         │ │
│  │ alpha_azi       [gaussian ▼]   [0.290]      [0.050]         │ │
│  └────────────────────────────────────────────────────────────┘ │
│                                                                  │
│  Distribution options for each parameter:                       │
│    • Fixed: Single value for all MMs                            │
│    • Gaussian: Mean + Sigma (normal distribution)               │
│    • Uniform: Min + Max (uniform distribution)                  │
│                                                                  │
│                    [Generate PSF Data]                           │
└─────────────────────────────────────────────────────────────────┘
```

**Features:**
- **Distribution Type dropdown**: Switches between gaussian/pseudo-voigt
- **Alpha parameters**: Only visible for pseudo-voigt
- **Dynamic fields**: 
  - "Fixed" shows 1 field (value)
  - "Gaussian" shows 2 fields (mean, sigma)
  - "Uniform" shows 2 fields (min, max)
- **Auto-clamping**: Alpha values clamped to [0, 1]

**Example Configurations:**

*All MMs Identical:*
- All parameters: "fixed"
- Result: Every MM has same PSF

*Varying Widths Only:*
- m_rad, m_azi: "fixed" = 0
- sigma_rad, sigma_azi: "gaussian" with mean=8, sigma=0.5
- Result: Same position, different widths

*Random Pseudo-Voigt Mix:*
- Distribution type: "pseudo-voigt"
- alpha_rad: "uniform", min=0.6, max=0.9
- alpha_azi: "uniform", min=0.2, max=0.4
- Result: Asymmetric PSFs with varying Gaussian/Lorentzian mix

**Screenshot location: Generate tab in Free mode with pseudo-voigt selected**

#### Step 4: Generate PSF Data

After configuring parameters (Standard or Free mode):

1. Click **"Generate PSF Data"** button
2. Progress indicator appears (for large datasets)
3. Success message: "PSF data generated for 145 MMs"
4. Data preview updates in Export tab

**What Happens:**
- Random values generated for each selected MM
- Alpha values clamped to [0, 1]
- Data stored in memory (not yet saved to Excel)
- Ready for export

#### Step 5: Other Data Types Tab

Generate additional perturbation data.

```
┌─────────────────────────────────────────────────────────────────┐
│  Other Data Types                                               │
├─────────────────────────────────────────────────────────────────┤
│  Select Data Type:                                              │
│    ◉ MM_alignement    ○ MM_gravity    ○ MM_thermal              │
│                                                                  │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │ Parameter       Distribution    Value 1      Value 2        │ │
│  ├────────────────────────────────────────────────────────────┤ │
│  │ dz [mm]         [gaussian ▼]   [0.000]      [0.010]         │ │
│  │ dx [mm]         [gaussian ▼]   [0.000]      [0.010]         │ │
│  │ dy [mm]         [gaussian ▼]   [0.000]      [0.010]         │ │
│  │ thetax [arcsec] [fixed    ▼]   [0.000]      (hidden)        │ │
│  │ thetay [arcsec] [fixed    ▼]   [0.000]      (hidden)        │ │
│  │ thetaz [arcsec] [fixed    ▼]   [0.000]      (hidden)        │ │
│  └────────────────────────────────────────────────────────────┘ │
│                                                                  │
│                   [Generate Selected Data]                       │
└─────────────────────────────────────────────────────────────────┘
```

**Available Types:**
- **MM_alignement**: Manufacturing/assembly errors (dz, dx, dy, rotations)
- **MM_gravity**: Gravity-induced deformations
- **MM_thermal**: Thermal expansion/contraction

**Usage:**
1. Select data type radio button
2. Configure parameters (same as PSF: fixed/gaussian/uniform)
3. Click **"Generate Selected Data"**
4. Repeat for other data types if needed
5. All generated data accumulates for export

**Screenshot location: Other Data Types tab with alignment errors configured**

#### Step 6: Export Tab

Preview and save all generated data.

```
┌─────────────────────────────────────────────────────────────────┐
│  Export                                                         │
├─────────────────────────────────────────────────────────────────┤
│  Export Mode:                                                   │
│    ◉ Update current file    ○ Create new file                  │
│                                                                  │
│  Output Path:                                                   │
│    [Distributions/TestDistribution.xlsx    ] [Browse...]       │
│                                                                  │
│  Data Preview:                                                  │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │ Sheet: MM_PSF                                               │ │
│  │ ┌──────────────────────────────────────────────────────────┤ │
│  │ │ MM# | m_rad | m_azi | sigma_rad | sigma_azi | distr |... │ │
│  │ │  1  | 0.000 | 5.234 |   6.512   |   6.128   | gauss |... │ │
│  │ │  2  | 0.000 | 4.891 |   6.701   |   6.445   | gauss |... │ │
│  │ │ ... (shows first 10 rows)                                │ │
│  │ └──────────────────────────────────────────────────────────┘ │
│  │                                                              │ │
│  │ Sheet: MM_alignement                                        │ │
│  │ ┌──────────────────────────────────────────────────────────┤ │
│  │ │ MM# |  dz   |  dx   |  dy   | thetax | thetay | thetaz  │ │
│  │ │  1  | 0.012 | 0.008 | -0.011|  0.000 |  0.000 |  0.000  │ │
│  │ │ ... (shows first 10 rows)                                │ │
│  │ └──────────────────────────────────────────────────────────┘ │
│  └────────────────────────────────────────────────────────────┘ │
│                                                                  │
│  ✓ PSF data ready                                               │
│  ✓ Alignment data ready                                         │
│                                                                  │
│                      [Export to Excel]                           │
│                                                                  │
│  Status: Ready to export                                        │
└─────────────────────────────────────────────────────────────────┘
```

**Export Modes:**

*Update Current File:*
- Overwrites sheets in the loaded Excel file
- Preserves MM configuration and other sheets
- **Preserves formulas and formatting**
- Safe: Creates backup before writing

*Create New File:*
- Saves to new filename
- Original file unchanged
- Auto-suggests name: `original_name_new.xlsx`

**Steps:**
1. Choose export mode
2. Verify output path (edit if needed)
3. Review data preview
4. Click **"Export to Excel"**
5. Success message: "Data exported successfully to Distributions/..."

**Screenshot location: Export tab with data preview and export button**

### Common Workflows

#### Workflow 1: Quick Standard Preset

**Goal:** Generate identical PSF for all MMs using preset

1. Load Excel file
2. MM Configuration: Select all (default)
3. PSF → Select: Choose "gaussian 10 arcsec HEW"
4. Click "Load Standard Distribution"
5. PSF → Generate: Click "Generate PSF Data"
6. Export: "Update current file", click "Export to Excel"

**Time:** ~30 seconds

#### Workflow 2: Variable PSF with Perturbations

**Goal:** Varying PSFs + alignment errors

1. Load Excel file
2. MM Configuration: Select desired MMs
3. PSF → Select: Standard mode, preset "gaussian 15 arcsec HEW 10%"
4. PSF → Generate: Generate PSF Data
5. Other Data Types: Select "MM_alignement"
   - dz: gaussian, mean=0, sigma=0.01
   - dx, dy: gaussian, mean=0, sigma=0.01
   - Click "Generate Selected Data"
6. Export: "Create new file", name "my_config.xlsx"

**Time:** ~2 minutes

#### Workflow 3: Custom Pseudo-Voigt

**Goal:** Full manual control with asymmetric alpha

1. Load Excel file
2. MM Configuration: Filter Row# = 1, select all in row
3. PSF → Select: Switch to "Free" mode
4. PSF → Generate:
   - Distribution Type: "pseudo-voigt"
   - m_rad: fixed = 0
   - m_azi: gaussian, mean=5, sigma=0.5
   - sigma_rad: gaussian, mean=8, sigma=0.8
   - sigma_azi: gaussian, mean=7, sigma=0.7
   - alpha_rad: fixed = 0.8 (heavy tails radially)
   - alpha_azi: fixed = 0.2 (sharp azimuthally)
5. Generate and export

**Time:** ~3 minutes

**Screenshot location: Example of asymmetric pseudo-voigt configuration**

### Keyboard Shortcuts and Tips

**General:**
- `Tab`: Navigate between fields
- `Enter`: Activate focused button
- `Esc`: Close dialogs

**Table Navigation:**
- Click column headers to sort
- Scroll with mouse wheel or trackpad
- Click checkboxes to select/deselect MMs

**Efficiency Tips:**
1. **Start with Standard presets** for speed
2. **Use filters** to quickly select MM subsets
3. **Preview before export** to catch errors
4. **Create new file first** when testing (preserve original)
5. **Generate all data types** before exporting (single export operation)
6. **Keep Excel closed** during export (avoids file locks)

### Troubleshooting

**GUI won't start:**
```bash
python3 --version  # Check Python ≥ 3.8
pip install -r requirements.txt  # Reinstall dependencies
```

**"No MM selected" error:**
- Go to MM Configuration tab
- Check at least one MM checkbox
- Click "Apply Selection"

**Parameters don't update:**
- In Standard mode, switch to Generate tab after loading preset
- In Free mode, ensure you clicked the Generate button

**Export fails:**
- Close Excel file if open
- Check file permissions
- Verify path exists: `ls -la Distributions/`

**Alpha values unexpected:**
- Check they're in [0, 1] range (auto-clamped)
- For pseudo-voigt only (hidden for gaussian)

### Advanced Features

**Batch Processing:**
Load different Excel files without restarting GUI - each Load clears previous configuration.

**Filter Combinations:**
- Row# = 1, MM# = All → All MMs in row 1
- Row# = All, MM# = 5 → MM #5 across all rows

**Data Validation:**
- Negative sigmas → Warning message
- Out-of-range alpha → Auto-clamped to [0, 1]
- Empty selection → Error before generation

**Excel Integration:**
- Formulas in Excel preserved during export
- Conditional formatting maintained
- Named ranges unchanged

---

## GUI Guide

### Workflow Example: Generate Pseudo-Voigt PSF Data

1. **Launch GUI**:
   ```bash
   python3 gui_distributions.py
   ```

2. **Load MM Configuration**:
   - Click "Load Excel File"
   - Select `Distributions/TestDistribution.xlsx` (or your file)
   - MM configuration automatically loads

3. **Select MMs** (MM Configuration tab):
   - All MMs are selected by default
   - Uncheck any MMs you want to exclude
   - Click "Apply Selection" (updates all data tabs)

4. **Configure PSF** (PSF → Generate tab):
   - **Option A: Use Standard Preset**:
     - Go to PSF → Select tab
     - Choose "Standard" mode
     - Select "pseudo-voigt 12 arcsec HEW 0.77/0.29" from dropdown
     - Click "Load Standard Distribution"
     - Parameters automatically filled
   - **Option B: Manual Configuration**:
     - Go to PSF → Select tab, choose "Free" mode
     - Go to PSF → Generate tab
     - Select "pseudo-voigt" from Distribution Type dropdown
     - **Alpha parameters appear**:
       - alpha_rad: gaussian, mean=0.77, sigma=0.05
       - alpha_azi: gaussian, mean=0.29, sigma=0.05
     - **Configure PSF parameters**:
       - m_rad: fixed = 0
       - m_azi: fixed = 5  
       - sigma_rad: gaussian, mean=8, sigma=0.5
       - sigma_azi: gaussian, mean=8, sigma=0.5
   - Click "Generate PSF Data"

5. **Export** (Export tab):
   - Mode: "New file"
   - Path: `Distributions/my_pseudo_voigt_data.xlsx`
   - Click "Export to Excel"

6. **Analyze** (using command line):
   ```bash
   python3 main.py -f Distributions/my_pseudo_voigt_data.xlsx
   ```

### Common Scenarios

#### All MMs with Fixed Parameters
1. Distribution type: "gaussian"
2. All parameters: "fixed" with desired values
3. Result: Identical PSF for all MMs

#### Varying Widths, Fixed Position
1. Distribution type: "gaussian" or "pseudo-voigt"
2. m_rad, m_azi: "fixed"
3. sigma_rad, sigma_azi: "gaussian" or "uniform"
4. Result: Same center, different widths per MM

#### Random Pseudo-Voigt Mix
1. Distribution type: "pseudo-voigt"
2. eta: "uniform", min=0.3, max=0.7
3. Configure other parameters as needed
4. Result: Each MM has different Gaussian/Lorentzian mix

## File Formats

### Input Excel File Structure

Required sheets:

1. **MM configuration**:
   ```
   | MM # | Row # | x_MM [m] | y_MM [m] | z_MM [m] | r_MM [m] |
   |------|-------|----------|----------|----------|----------|
   |  1   |   1   |  0.123   |  0.456   |  7.89    |  0.15    |
   |  2   |   1   |  0.234   |  0.567   |  7.89    |  0.15    |
   ```

2. **MM_PSF** (generated by GUI):
   ```
   | MM # | m_rad [arcsec] | m_azi [arcsec] | sigma_rad [arcsec] | sigma_azi [arcsec] | distribution | alpha_rad | alpha_azi |
   |------|----------------|----------------|--------------------|--------------------|--------------|-----------|----------|
   |  1   |      0         |       5        |        8.2         |        7.8         | pseudo-voigt |   0.77    |   0.29   |
   |  2   |      0         |       5        |        8.5         |        8.1         | pseudo-voigt |   0.75    |   0.31   |
   ```

3. **Alignment** (optional):
   ```
   | Position # | d_align_rad [µm] | d_align_azi [µm] | d_align_z [µm] | d_align_rotz [arcsec] |
   ```

4. **Gravity offload** (optional):
   ```
   | Position # | d_grav_x [µm] | d_grav_y [µm] | d_grav_z [µm] | d_grav_rotz [arcsec] |
   ```

5. **Thermal** (optional):
   ```
   | Position # | d_therm_x [µm] | d_therm_y [µm] | d_therm_z [µm] | d_therm_rotz [arcsec] |
   ```

**Axis interpretation (for Gravity/Thermal and related terms):**
- `x`, `y`: shifts in the x and y directions (lateral displacements)
- `z`: focus change (axial displacement)
- `rotz`: rotation about the optical axis
- `rad`: shift in the radial direction (positive is towards higher radius / outward)
- `azi`: shift along the azimuthal direction (positive azimuth shifts are clockwise with respect to the radial vector)

**Sign conventions:**
- Positive `rotz` rotations introduce positive shifts in the azimuthal direction.

### Output Figures

Figures are saved in `Figures/` folder with timestamps:

- **PSF Plot**: `Figures/E2E_PSF_YYYYMMDD_HHMMSS.png`
- **Encircled Energy**: `Figures/Encircled_Energy_YYYYMMDD_HHMMSS.png`

Format: PNG, 300 DPI, high quality for publications.

## Tips and Best Practices

### GUI Tips
1. **Always apply selection**: After changing MM selection in the configuration tab, click "Apply Selection"
2. **Preview before export**: Use the Export tab preview to verify data
3. **Hidden fields**: When distribution is "fixed", the second parameter field is hidden (not just disabled)
4. **Alpha visibility**: Alpha controls only appear for pseudo-voigt distribution type (in Free mode)
5. **Standard vs Free mode**: Use Standard mode for quick preset loading, Free mode for full control
6. **Variable preset derivation**: Variable gaussian presets automatically calculate mean sigma from HEW parsing
7. **Alpha clamping**: All alpha values automatically clamped to valid range [0, 1]

### Analysis Tips
1. **Start with coarse mode**: Use `--mode coarse` for quick previews
2. **Normalize for comparison**: Use `--normalize` to compare different configurations
3. **Save important plots**: Use keyboard shortcuts or context menu to export specific plots
4. **Batch processing**: Use shell scripts to process multiple files

### Parameter Selection
1. **Sigma values**: Typically 5-10 arcsec for good optics
2. **Alpha variation**: Keep sigma small (0.05-0.15) to avoid extreme values
3. **Uniform ranges**: Use tight ranges [0.4, 0.6] for controlled variation
4. **Fixed baseline**: Start with all fixed parameters, then vary one at a time
5. **HEW to sigma conversion**: Use formula σ = HEW / (2√(2ln2)) or standard presets for automatic conversion
6. **Optimization**: Run with `--optimize` if MMs have varied PSF parameters to find best row-wise arrangement

## Troubleshooting

### Common Issues

**GUI doesn't start:**
```bash
# Check Python version
python3 --version  # Must be 3.8+

# Check dependencies
pip install --upgrade -r requirements.txt
```

**File not found errors:**
```bash
# Ensure you're running from the project root
cd "/path/to/NewAthenaE2EPSF_v2"
python3 gui_distributions.py
```

**Figures not saving:**
```bash
# Check Figures directory exists
ls -la Figures/

# Create if missing
mkdir -p Figures
```

**Excel file errors:**
- Ensure file has "MM configuration" sheet
- Check column names match exactly (case-sensitive)
- Verify Excel file is not open in another program

## Version Information

**Version**: 3.0  
**Date**: 2026-01-28  
**Python**: 3.8+

### Notable changes in v3.0

- Deterministic per‑MM sampling for presets: CSV and Excel generation now expand presets into numeric per‑MM `sigma`/`alpha` values, reproducible for a given filename/preset/index.
- CSV fast-path parity with Excel: multi-sheet CSVs are materialized and handled equivalently to `.xlsx` inputs.
- Removed legacy in‑memory/pickle flows and debug helper tools; generator no longer relies on persisted `.pkl` inputs and temporary persisted input folders were cleaned.
- Renamed directory `sensivitiy` → `sensitivity` and updated code references and tests accordingly.
- Replaced ad-hoc debug prints with proper `logging`; added `pytest.ini` and unit tests to validate CSV parsing and distribution parsing (tests run locally).
- Documentation updates: `QUICKSTART.txt`, `README.md`, and `requirements.txt` updated (includes tkinter install notes for macOS/Linux/Windows).
- Release prep: created release package and pushed tag `v3` to remote.

### Notable changes in v2.0

- Pseudo-Voigt support with independent `alpha_rad` / `alpha_azi`
- Standard preset loading from the workbook (`MM_PSF` preset table starting at cell K1)
- Row-wise MM# optimizer + deterministic placement seeds (`cross`, `x_axis`, `elliptical`)
- Rotation-invariant HEW evaluation on a polar grid
- Figure exports saved into `Figures/` with timestamped filenames

### Runtime mode note

The `--mode {coarse,fine}` flag trades speed vs sampling density for both plotting and optimization. Start with `coarse` to validate inputs and workflow, then switch to `fine` for final runs if needed.

---

**License**: Internal Use
