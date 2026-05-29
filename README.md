# NewAthenaE2EPSF — PSF Analysis Toolkit

[![Python CI](https://github.com/Astro-cl/NewAthena_E2E_PSF/actions/workflows/python-ci.yml/badge.svg)](https://github.com/Astro-cl/NewAthena_E2E_PSF/actions/workflows/python-ci.yml)

End-to-end PSF simulation and analysis tool for the NewAthena X-ray telescope mirror module assembly. Reads a configured Excel workbook describing mirror module (MM) PSF parameters, alignment, gravity, and thermal perturbations, then computes the aggregate E2E PSF, encircled energy function, HEW metrics, and radial profile fits.

> **Note:** Documentation, comments and unit tests were written with AI assistance.

---

## Contents

- [Installation](#installation)
- [Workflow Overview](#workflow-overview)
- [Command-Line Interface](#command-line-interface)
- [GUI Application](#gui-application)
- [Distribution Types](#distribution-types)
- [Perturbation Model](#perturbation-model)
- [PSF Fitting Pipeline](#psf-fitting-pipeline)
- [Row-Wise MM Optimizer](#row-wise-mm-optimizer)
- [Batch Combinations](#batch-combinations)
- [Sensitivity Pipeline](#sensitivity-pipeline)
- [Export Packages](#export-packages)
- [A_eff and Vignetting](#aeff-and-vignetting)
- [Excel File Format](#excel-file-format)
- [Directory Structure](#directory-structure)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [Author](#author)
- [Release History](#release-history)

---

## Installation

### Requirements

- Python 3.8 or higher
- pip

### Setup

```bash
# Create and activate a virtual environment (recommended)
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

Required packages: `numpy`, `pandas`, `matplotlib`, `openpyxl`, `scipy`, `lmfit`.

### Tkinter (GUI)

The GUI uses `tkinter`. This is included with CPython but requires the underlying Tcl/Tk libraries on the host. It is **not** installed via pip.

| Platform | Installation |
|----------|-------------|
| **macOS** | Install Python from python.org (bundles Tcl/Tk), or: `brew install tcl-tk` |
| **Debian/Ubuntu** | `sudo apt install python3-tk` |
| **Fedora/RHEL** | `sudo dnf install python3-tkinter` |
| **Arch Linux** | `sudo pacman -S tk` |
| **Windows** | Use the official installer from python.org |

---

## Workflow Overview

The typical workflow is:

1. **GUI** (`gui_distributions.py`) — configure MM PSF parameters and export a populated Excel workbook
2. **CLI** (`main.py`) — read the workbook and run the full PSF analysis, metric computation, and export

The GUI is a configuration helper; `main.py` is the analysis engine. Once the workbook is prepared, all analysis is driven from the CLI.

```
GUI (gui_distributions.py)
  └─ Export Excel workbook  →  CLI (main.py)
                                  └─ Figures, FITS, export packages
```

---

## Command-Line Interface

```bash
python3 main.py -f Distributions/your_file.xlsx [options]
```

### Options

| Option | Description |
|--------|-------------|
| `-f / --file` | Path to input Excel workbook |
| `--mode {coarse,fine}` | Computation resolution (default: coarse) |
| `--output PATH` | Save combined figure to PATH and exit without opening a window |
| `--normalize` | Normalize PSF for plot comparison |
| `--export-package` | Package figures, FITS, and workbook into `Exports/<TIMESTAMP>/` |
| `--single-config N` | Only run configuration index N (for debugging) |
| `--optimize` | Enable row-wise MM# assignment optimization |
| `--placement STRATEGY` | Placement strategy seed for optimizer (cross, x_axis, elliptical) |
| `--batch-combinations FILE` | Automated multi-configuration batch run |
| `--log-dz` | Log per-MM dz projections |

### Interactive Mode

Running without `--output` or `--export-package` opens the **E2E PSF Viewer - MM Selector** window. A collapsible **Row -> Petal -> MM** checkbox tree on the left lets you select any subset of mirror modules; clicking **Update E2E & EEF** re-renders the PSF and encircled energy function for the selection. A **HEW Contribution Ranking** panel at the bottom ranks checked MMs by individual HEW contribution (leave-one-out analysis), displayed in a sortable table with colour coding. A **Map** button shows a 2D scatter plot of MM positions colour-coded by delta-HEW.

**Keyboard / context menu shortcuts (interactive window):**
- `p` — export PSF PNG to `Figures/`
- `e` — export EEF PNG to `Figures/`
- `f` — export PSF to FITS in `CustomPSFs/`
- `4` or `c` — export EEF CSV to `CustomPSFs/`
- Right-click on a plot for all of the above plus fit parameters CSV

### Examples

```bash
# Interactive analysis
python3 main.py -f Distributions/my_config.xlsx

# Save figure without opening a window
python3 main.py -f Distributions/my_config.xlsx --output Figures/result.png

# Export package (coarse)
python3 main.py -f Distributions/my_config.xlsx --export-package --mode coarse

# Export package (fine resolution, compute-intensive)
python3 main.py -f Distributions/my_config.xlsx --export-package --mode fine

# Debug: run only configuration index 1
python3 main.py -f Distributions/my_config.xlsx --export-package --single-config 1

# Batch: sweep over off-axis / energy / defocus combinations
python3 main.py -f Distributions/my_config.xlsx --batch-combinations combinations.xlsx

# Optimize MM# assignments within rows
python3 main.py -f Distributions/my_config.xlsx --optimize
```

**Background / unattended runs:**

```bash
nohup python3 main.py -f Distributions/my_config.xlsx --export-package --mode fine &
# or with screen:
screen -S e2e_run
python3 main.py -f Distributions/my_config.xlsx --export-package --mode fine
# detach: Ctrl-A D
```

---

## GUI Application

Launch with:

```bash
python3 gui_distributions.py
```

> **macOS trackpad note:** Tap-to-click is supported. Clicking anywhere on a combobox field opens the dropdown list.

The GUI has four main tabs: **MM Configuration**, **PSF**, **Other Data Types**, and **Export**.

### MM Configuration Tab

Displays all MMs from the loaded Excel file in a scrollable, sortable table.

- **Checkboxes** — select which MMs to include in PSF generation
- **Select/Deselect All** — toggle all at once
- **Filters** — filter by Row # or MM # (e.g. Row=1 shows all MMs in row 1)
- **Apply Selection** — must be clicked to confirm changes

### PSF Tab

Two modes are available via radio buttons: **Standard** and **Free**.

#### Standard Mode

Loads predefined PSF parameter sets from the `MM_PSF` sheet preset table (starting at cell **M1** of the workbook).

1. Select **Standard** radio button
2. Choose a preset from the dropdown **"Load from Excel (M1):"**
3. Click **Load Standard Distribution**
4. Switch to the **Generate** sub-tab and click **Generate PSF Data**

In Standard mode, parameters are read-only. Preset types:
- **Fixed presets** — all MMs receive identical values
- **Variable presets** (contain `%`) — parameters vary per MM using a Gaussian distribution; sigma_mean is derived from HEW, sigma_sigma = sigma_mean x percent
- **Pseudo-Voigt presets** — include `alpha_rad`/`alpha_azi` values

#### Free Mode

Full manual control over all parameters.

1. Select **Free** radio button
2. Choose distribution type (gaussian / pseudo-voigt)
3. For each parameter, set distribution type and values:
   - **Fixed**: single value
   - **Gaussian**: mean and sigma
   - **Uniform**: min and max
4. Alpha parameters appear only for pseudo-voigt; values are clamped to [0, 1]
5. Click **Generate PSF Data**

#### Preset Table Format (MM_PSF sheet, cell M1)

Each row defines one preset. Columns (left to right):

| Column | Content |
|--------|---------|
| 1 | Preset name |
| 2 | `sigma_rad` spec |
| 3 | `sigma_azi` spec |
| 4 | `alpha_rad` spec (pseudo-voigt; use `-` or blank for Gaussian) |
| 5 | `alpha_azi` spec |

Spec cell formats: plain number (fixed), `gaussian(mean, sigma)`, `uniform(min, max)`. All values in **arcsec**.

**Reference sigma values for common presets:**

| Preset | sigma_rad | sigma_azi |
|--------|-----------|-----------|
| Symmetric Gaussian, 4.3 arcsec HEW | 1.826058874 | 1.826058874 |
| Symmetric Gaussian, 8.0 arcsec HEW | 3.397323952 | 3.397323952 |
| Asymmetric Gaussian, 4.3 arcsec HEW, ratio 4:1 | 2.963740901 | 0.740935225 |
| Asymmetric Gaussian, 8.0 arcsec HEW, ratio 7:1 | 5.797205089 | 0.828172156 |

> Note: For asymmetric Gaussians, `sigma = HEW / 2.355` does **not** apply. Use sigma values that reproduce the correct 50% encircled-energy diameter (see table above).

**Common pitfalls:**
- Reload the workbook in the GUI after editing the preset table in Excel
- Use dot decimals (`3.397`, not `3,397`)
- Alpha columns for Gaussian presets should be `-` or blank
- Do not put units in cells -- values are always in arcsec

### Other Data Types Tab

Generate perturbation data alongside the PSF:
- **MM_alignment** — manufacturing/assembly errors (dx, dy, dz, rotations)
- **MM_gravity** — gravity-induced deformations
- **MM_thermal** — thermal expansion/contraction

Each type uses the same fixed/Gaussian/uniform parameter scheme. Click **Generate Selected Data** for each type needed.

### Export Tab

**Export Modes:**
- **Update Current File** — overwrites generated sheets; preserves formulas, formatting, and other sheets; creates a backup before writing
- **Create New File** — saves to a new filename; auto-suggests `original_name_new.xlsx`

**A_eff export behavior:**
- GUI export evaluates standard A_eff presets per-MM and writes numeric values into column B of the `A_eff` sheet; column C is cleared
- CLI (`main.py`) preserves legacy behavior: adjusted/derived A_eff is written to column C

**Vignetting export:**
- The A_eff tab includes an **"Apply vignetting factors when exporting"** checkbox
- When enabled with a standard preset, the export copies the chosen preset column from the `Vignetting rotazi` and `Vignetting rotrad` sheets into column B

**Precomputing A_eff values (for formula-heavy workbooks):**

```bash
python3 tools/compute_aeff_values.py path/to/your_workbook.xlsx
```

This writes numeric A_eff into the `A_eff` sheet so downstream runs do not depend on Excel formula evaluation.

---

## Distribution Types

### Gaussian

Standard rotated 2D Gaussian:

```
dx = x - mu_x;  dy = y - mu_y
c = cos(theta); s = sin(theta)

a     = c^2/sx^2 + s^2/sy^2
b     = s*c*(1/sx^2 - 1/sy^2)
cc    = s^2/sx^2 + c^2/sy^2

G(x,y) = A * coeff * exp(-0.5*(a*dx^2 + 2*b*dx*dy + cc*dy^2))
```

where `coeff = 1/(2*pi*sx*sy)` when `normalize=True`, else `1`.

Parameters: `mux`, `muy` (center, m); `sigmax`, `sigmay` (principal sigmas, m); `theta` (rotation, rad); `amplitude`.

### Pseudo-Voigt

Separable product of two 1D Pseudo-Voigt functions:

```
PV(azi, rad) = PV_azi(azi/sigma_azi) * PV_rad(rad/sigma_rad)
PV_1D(u) = (1 - alpha) * G(u) + alpha * L(u)
```

where `G(u) = (1/sqrt(2*pi)) * exp(-u^2/2)` and `L(u) = (1/pi)/(1 + u^2)`.

- `alpha = 0` -- pure Gaussian (sharp core)
- `alpha = 0.5` -- balanced mix
- `alpha = 1` -- pure Lorentzian (heavy tails)
- Independent `alpha_rad` and `alpha_azi` are supported for directional asymmetry

### Modified Pseudo-Voigt (Aggregate Radial Fit)

Used to fit the azimuthal-average radial intensity profile of the aggregated PSF:

```
G(r; Gc) = exp(-4*ln2*(r/Gc)^2)
a         = 2^(1/beta) - 1
C(r; Gw) = [1 + a*(2r/Gw)^2]^(-beta)
I(r)      = A * [(1-eta)*G + eta*scalar*C] / [(1-eta) + eta*scalar]
```

| Parameter | Description |
|-----------|-------------|
| `A` | Amplitude |
| `Gamma_core` | Core FWHM (arcsec) |
| `Gamma_wing` | Wing width (arcsec) |
| `eta` | Core-to-wing mixing fraction |
| `beta` | Wing shape exponent |
| `scalar` | Wing amplitude scale |

### Pearson Type IV

1D profile fitted to the radial aggregate:

```
P(u) proportional to (1 + (u/sigma)^2)^(-m) * exp(nu * atan(u/sigma)),   u = x - mu
```

Parameters: `mu` (center), `sigma` (scale), `m` (tail shape, `m > 1`), `nu` (skew; `nu = 0` is symmetric). Skipped in coarse mode for performance.

### Unit Conventions

- **Arcsec to meters:** `1 arcsec = 12 * pi / 180 / 3600 m` (focal length f = 12 m)
- **HEW to sigma (symmetric Gaussian):** `sigma = HEW / (2*sqrt(2*ln2)) approx HEW / 2.3548`
- All public APIs use **meters**; the loader converts arcsec on read

---

## Perturbation Model

Perturbation contributions from four sources are combined additively to compute centroid offsets for each MM:

| Source | Sheet name | Columns |
|--------|-----------|---------|
| Alignment | `Alignment` | `d_align_x/y/z [um]`, `d_align_rotx/y/z [arcsec]` |
| Gravity offload | `Gravity offload` | `d_grav_x/y/z [um]`, `d_grav_rotx/y/z [arcsec]` |
| Thermal | `Thermal` | `d_therm_x/y/z [um]`, `d_therm_rotx/y/z [arcsec]` |
| Extra (off-axis / defocus) | `Extra PSF shifts` | `d_extra_rotx/y [arcsec]`, `d_extra_z [um]` |

### d_z to Centroid Shift

Axial displacement `d_z` (sum of all four sources) projects to a focal-plane centroid shift:

```
dm_x = d_z_total * x_MM / (f - z_MM)
dm_y = d_z_total * y_MM / (f - z_MM)
```

where `f = 12 m` is the focal length and `x_MM`, `y_MM`, `z_MM` are the MM geometric coordinates from the `MM configuration` sheet.

### Off-Axis Pointing

An off-axis angle specified in arcminutes is decomposed equally into X and Y rotation components and written to the `Extra PSF shifts` sheet:

```
d_extra_rotx [arcsec] = offaxis_arcmin * 60 / sqrt(2)
d_extra_roty [arcsec] = offaxis_arcmin * 60 / sqrt(2)
```

These are added to the total rotation in `compute_total_rot_polar()` alongside alignment, gravity, and thermal contributions.

### Defocus

Defocus (in mm) is converted to um and written to `d_extra_z [um]` in the `Extra PSF shifts` sheet, then loaded as metres (* 1e-6) and added to the total `d_z` before projection.

Unit chain: `defocus [mm]` -> `* 1e3` -> `[um]` (sheet) -> `* 1e-6` -> `[m]` (code)

### HEW Sigma Broadening

Two optional sheets provide per-position HEW degradation as a function of off-axis angle and energy:

- `MM HEW degradation rotazi` -- azimuthal HEW degradation, broadens `sigma_azi`
- `MM HEW degradation rotrad` -- radial HEW degradation, broadens `sigma_rad`

Each sheet has:
- **Columns A-C**: per-position output (`Position #`, `HEW degradation (arcsec)`, `Selected energy [keV]` in C2)
- **Columns H-K**: lookup table (`Row #`, `angle [arcmin]`, `energy [keV]`, `HEW [arcsec]`)

Processing: MM rotation angle (arcsec) -> `np.interp` against table -> HEW degradation per position -> sigma broadening in quadrature:

```
sigma_extra = HEW_deg / (2*sqrt(2*ln2)) * arcsec_to_m
sigma_new   = sqrt(sigma_base^2 + sigma_extra^2)
```

Results written to `Extra PSF degradations` sheet (cols B/C: sigma_extra_rad/azi in arcsec) and MM_PSF columns I/J (final degraded sigma).

### Energy-Dependent Sigma Scaling

After angle-based broadening, a multiplicative energy scaling factor is applied from the `MM HEW degradation energy` sheet:

| Column A | Column B |
|----------|----------|
| Energy [keV] | Sigma scaling factor |

The factor at the selected energy is linearly interpolated (boundary-clamped) and applied to both `sigma_rad` and `sigma_azi` for all MMs:

```
sigma_final = sigma_broadened * f(E)
```

The selected energy comes from cell `A_eff!D2` (propagated to other sheets in batch mode). This step runs even when no HEW degradation sheets are present.

Combined model: sigma_final = sqrt(sigma_base^2 + sigma_extra^2) * f(E)

### Defocus PSF Shape Broadening

Axial defocus also broadens the per-MM PSF shape. The PSF size (6*sigma) grows linearly with dz from its initial value at best focus to the full MM physical dimension at `dz = 12 m`:

```
sigma_rad_adj = sigma_rad_init + (MM_height - 6*sigma_rad_init) / 12 * dz / 6
sigma_azi_adj = sigma_azi_init + (MM_width  - 6*sigma_azi_init) / 12 * dz / 6
```

`MM_height` and `MM_width` are read from columns I and J of the `MM configuration` sheet. `dz` is the signed total axial displacement in metres. Adjusted values are written back to MM_PSF columns I/J.

### VLOOKUP Resolver

When `main.py` saves and reloads a workbook, openpyxl strips cached Excel formula values. If MM_PSF columns D/E (base sigma) contain `VLOOKUP` formulas with no cached values, a Python fallback resolves them: MM# -> Row# (via `MM configuration` column C) -> (sigma_rad, sigma_azi) from the preset table at MM_PSF rows M30:Q45. Resolved values are written back as plain numbers.

---

## PSF Fitting Pipeline

### Stage 1 -- PSF Aggregation

Individual MM PSFs are co-added on a shared 2D Cartesian grid:

```
Z(x, y) = sum_i w_i * PSF_i(x - mu_xi, y - mu_yi)
```

where `w_i` is the A_eff weight and `(mu_xi, mu_yi)` is the centroid after perturbation projection.

### Stage 2 -- Best-Focus Minimisation

The focal-plane centre `(cx, cy)` minimising the HEW is found by a discrete gradient search starting from the nominal centre (1 um steps, up to 30 iterations, multiple candidate starts).

### Stage 3 -- Radial Profile

The azimuthally averaged radial intensity is computed on a polar grid centred at `(cx, cy)`:

```
E(r_k) = sum_j Z(cx + r_k*cos(t_j), cy + r_k*sin(t_j)) * r_k * dt
Phi(r) = cumsum_k E(r_k) * dr      (cumulative radial energy)
I(r)   = E(r) / (2*pi*r)           (mean radial intensity)
```

Default grid: N_r = 400, N_theta = 360. The grid expands automatically (up to 3x, factor 1.5 per iteration) if less than 99.95% of total energy is enclosed.

### Stage 4 -- EEF and HEW

```
EEF(d) = Phi(d/2) / Phi(r_max)
```

Radii enclosing 50%, 80%, 90% of energy are located by cubic interpolation:

| Metric | Definition |
|--------|-----------|
| `HEW` | 2 * r_50% (half-energy diameter, arcsec) |
| `EEF-80` | 2 * r_80% |
| `EEF-90` | 2 * r_90% |

### Stage 5 -- Modified Pseudo-Voigt Radial Fit

Fitted in two stages:
1. **Seed**: Gaussian fit to the core region gives initial `A` and `Gamma_core`
2. **Full fit**: Modified pseudo-Voigt minimising EEF residuals via `scipy.optimize.least_squares` with `soft_l1` loss and 36 multi-start perturbations (+/-12% random offsets)

Parameter bounds:

| Parameter | Lower | Upper |
|-----------|-------|-------|
| `A` | 0.5 * A0 | inf |
| `Gamma_core` | 0.2 * Gc0 | 3 * Gc0 |
| `Gamma_wing` | 1.2 * Gc0 | 25 * Gc0 |
| `eta` | 0.05 | 0.50 |
| `beta` | 1.0 | 5.0 |
| `scalar` | 0.2 | 12.0 |

### Stage 6 -- Pearson Type IV Fit

Fitted with an EEF-objective least-squares using 36 multi-starts (lmfit). Bounds: `m in [1, 8]`, `nu in [-2, 2]`. Skipped in coarse mode.

### Stage 7 -- FWHM from 1D Marginals

```
prof_x(x) = integral Z(x,y) dy
FWHM_x    = x_right_half_max - x_left_half_max
```

### Stage 8 -- Directional HEW from Marginals

Minimum-width interval of each 1D marginal enclosing exactly 50% of total marginal energy.

### Output Metrics

| Metric | Description |
|--------|-------------|
| `HEW` | Half-energy diameter (arcsec), rotation-invariant |
| `EEF-80` / `EEF-90` | Diameters enclosing 80%/90% of energy (arcsec) |
| `FWHM_x` / `FWHM_y` | Half-maximum widths from 1D marginals (arcsec) |
| `HEW_x` / `HEW_y` | Marginal half-energy widths (arcsec) |
| `Gamma_core` / `Gamma_wing` | Modified PV core and wing FWHM parameters (arcsec) |
| `eta` / `beta` / `scalar` | Modified PV mixing and shape parameters |

---

## Row-Wise MM Optimizer

Finds the best MM# assignment within each physical row of the mirror, keeping MM positions fixed, to minimise aggregate HEW.

```bash
python3 main.py -f Distributions/input.xlsx --optimize
```

Outputs: `Distributions/my_data_optimised.xlsx` and `Figures/E2E_PSF_*_optimised_*.png`.

**Placement strategies:**

| Strategy | Behaviour |
|----------|-----------|
| `elliptical` | Best MMs near x-axis slots, worst near y-axis (default seed) |
| `cross` | Best MMs on +/-x/y cross pattern to reduce anisotropy |
| `x_axis` | Best MMs near +/-x, alternating above/below to avoid clustering |
| `best_center` | Best MMs closest to optical axis |
| `worst_center` | Worst MMs at centre, best at edges |
| `alternating` | Alternates best/worst from centre outward |
| `random` | Random baseline |

Use `--placement STRATEGY` to apply a placement independently of `--optimize`.

---

## Batch Combinations

The `--batch-combinations` flag automates multi-configuration runs over off-axis angles, energies, and defocus values defined in a spreadsheet.

```bash
python3 main.py -f Distributions/base.xlsx --batch-combinations combinations.xlsx --mode coarse
```

### Combinations File Format

One configuration per row:

| Column | Content | Unit |
|--------|---------|------|
| A | Row ID | -- |
| B | Configuration name (used in output naming) | -- |
| C | Off-axis angle | arcmin |
| D | Energy | keV |
| E | Defocus | mm |
| F | Run mode (`coarse` / `fine`, optional) | -- |

### Per-Configuration Processing

For each row the tool:
1. Copies the base workbook with prefix-based naming
2. Writes off-axis rotations and defocus to the `Extra PSF shifts` sheet
3. Sets the energy in vignetting and HEW degradation sheets
4. Runs the full PSF analysis pipeline in headless mode (matplotlib Agg backend)
5. Creates an export package and ZIP under `Exports/Export_<input>_<timestamp>/`

After all configurations, an aggregated workbook is written containing per-configuration HEW, EEF, A_eff loss, and fit parameters.

---

## Sensitivity Pipeline

A sweep runner under `sensitivity/` generates per-combination input workbooks and optionally executes them.

```bash
# Generate input workbooks only
python3 sensitivity/sensitivity_run.py --generate-only --baseline Distributions/TestDistribution.xlsx

# Generate and run all jobs
python3 sensitivity/sensitivity_run.py --baseline Distributions/TestDistribution.xlsx
```

- Generated workbooks are written to `sensitivity/input/` (newest 100 kept)
- Per-job partial results are appended to `sensitivity/results/sensitivity_run_partial.csv` as jobs complete
- Final consolidated results are written to `sensitivity/results/sensitivity_run_results.xlsx`

**MM_PSF edits during sensitivity runs:**
- Only per-MM input columns B-H are modified; right-hand template columns are preserved
- For non-pseudo-voigt presets, `alpha_rad` and `alpha_azi` are set to `-`
- When a combo sets `Alignment=0`, `Gravity=0`, or `Thermal=0`, the corresponding perturbation columns B-G in the respective sheet are zeroed

---

## Export Packages

Packages written to `Exports/<TIMESTAMP>/` include:
- Packaged workbook (authoritative source for `Aeff_sum_orig` / `Aeff_sum_mod`)
- Figures in `Figures/`
- FITS files in `CustomPSFs/`
- Aggregated Excel `E2E_EEF_and_fitparams_*.xlsx`

**Verification after export:**
- PNGs: coarse = 320x320 px, fine = 2062x2062 px
- FITS dimensions match requested pixel size
- Aggregated Excel present in the package

### Repairing Existing Packages

| Script | Purpose |
|--------|---------|
| `.inspect_export.py` | Inspect a package and list mismatches between workbook and aggregated sums |
| `.diagnose_aeff.py` | Identify formula-only A_eff columns with missing cached values |
| `.patch_fitparams.py` | Patch per-config `fitparams_aeffloss.xlsx` files inside a package |
| `.patch_aggregated.py` | Repair aggregated workbook rows that used the wrong A_eff source |

```bash
python3 .inspect_export.py Exports/20260416_124558
python3 .patch_fitparams.py Exports/20260416_124558 --dry-run
python3 .patch_fitparams.py Exports/20260416_124558
```

---

## A_eff and Vignetting

### A_eff Sheet

- Column A: `MM #`
- Column B: weight (numeric; missing or non-numeric values cause analysis to fail)

**GUI export:** evaluates standard presets per-MM and writes numeric values to column B; clears column C.
**CLI:** writes adjusted/derived A_eff to column C (legacy behavior preserved).

**Supported A_eff preset expressions:**
- `J` -- copy value from column letter J
- `gaussian(J, sigma)` -- sample per-MM around column value
- `J+gaussian(0, 20%*J)` -- additive Gaussian noise

**Percent-variable presets** (e.g. `Variable 10% 1 keV`) are internally synthesized to explicit Gaussian forms for deterministic per-index sampling. Energy tokens like `1 keV` in preset names are parsed and written into cell `C2` of vignetting sheets at export time.

### Vignetting

Two vignetting sheet formats are supported:
- **Two-column** (delta -> factor)
- **Per-position columns**

Vignetting is applied after A_eff initialisation with explicit bookkeeping: `aeff_base`, `aeff_adjusted`, `aeff_vig_factor`. All `np.interp()` calls use `abs(rotation_value)` -- tables use non-negative delta values and symmetry is assumed.

The GUI's **"Apply vignetting factors when exporting"** checkbox copies the chosen preset column from `Vignetting rotazi` and `Vignetting rotrad` into column B during export.

---

## Excel File Format

### Required Input Sheets

**`MM configuration`**

| Column | Header | Description |
|--------|--------|-------------|
| A | `MM #` | Mirror module position number |
| B | `Row #` | Row assignment |
| C | `x_MM [m]` | X position (m) |
| D | `y_MM [m]` | Y position (m) |
| E | `z_MM [m]` | Z position (m) |
| F | `r_MM [m]` | Radial position (m) |
| I | `MM_height [m]` | MM height (used for defocus broadening) |
| J | `MM_width [m]` | MM width (used for defocus broadening) |

**`MM_PSF`** (generated by GUI)

| Column | Header | Description |
|--------|--------|-------------|
| A | `MM #` | Position number |
| B | `m_rad` | Radial shape parameter |
| C | `m_azi` | Azimuthal shape parameter |
| D | `sigma_rad [arcsec]` | Radial sigma (base) |
| E | `sigma_azi [arcsec]` | Azimuthal sigma (base) |
| F | `distribution` | `gaussian` or `pseudo-voigt` |
| G | `alpha_rad` | Radial mixing parameter [0,1] |
| H | `alpha_azi` | Azimuthal mixing parameter [0,1] |
| I | `sigma_rad_deg [arcsec]` | Final degraded radial sigma (written by CLI) |
| J | `sigma_azi_deg [arcsec]` | Final degraded azimuthal sigma (written by CLI) |

**`A_eff`**

| Column | Description |
|--------|-------------|
| A | `MM #` |
| B | Weight (numeric) |

**Optional perturbation sheets:** `Alignment`, `Gravity offload`, `Thermal`, `Extra PSF shifts`, `MM HEW degradation rotazi`, `MM HEW degradation rotrad`, `MM HEW degradation energy`.

### Axis Conventions

- `x`, `y` -- lateral displacements
- `z` -- axial / focus displacement (positive = towards detector)
- `rotz` -- rotation about the optical axis (positive introduces positive azimuthal shift)
- `rad` -- radial direction (positive = outward)
- `azi` -- azimuthal direction (positive = clockwise relative to radial vector)

### Output Files

| File | Description |
|------|-------------|
| `Figures/E2E_PSF_YYYYMMDD_HHMMSS.png` | PSF plot at 300 DPI |
| `Figures/Encircled_Energy_YYYYMMDD_HHMMSS.png` | EEF plot at 300 DPI |
| `Figures/E2E_fit.png` | Aggregated radial profile fit |
| `CustomPSFs/E2E_EEF_YYYYMMDD_HHMMSS.csv` | EEF data as CSV |
| `CustomPSFs/E2E_aggregated_*.fits` | PSF matrix in FITS format |

---

## Directory Structure

```
NewAthenaE2EPSF/
├── README.md                        # This file
├── requirements.txt                 # Python dependencies
├── main.py                          # CLI analyzer and PSF engine
├── gui_distributions.py             # Interactive GUI
├── distributions_rotated.py         # Core distribution functions
├── optimize_mm_rows.py              # MM optimizer and placement strategies
├── Distributions/                   # Input Excel workbooks
├── Figures/                         # Exported plots
├── CustomPSFs/                      # FITS and EEF CSV outputs
├── Exports/                         # Export packages
├── sensitivity/                     # Sensitivity pipeline
├── tools/                           # Helper and diagnostic scripts
└── tests/                           # Unit and integration tests
```

**Core module roles:**
- `main.py` -- CLI entry points, Excel I/O, perturbation application, PSF aggregation and fitting, export
- `gui_distributions.py` -- Tkinter GUI for generating and exporting per-MM distribution configurations
- `distributions_rotated.py` -- Gaussian, Pseudo-Voigt, Pearson4, King distribution functions
- `optimize_mm_rows.py` -- row-wise MM# optimizer, placement strategies, Excel write helpers

---

## Troubleshooting

**GUI won't start**

```bash
python3 --version            # Must be 3.8+
pip install -r requirements.txt
python3 -c "import tkinter"  # Should print nothing if OK
```

If tkinter is missing, install the OS package (see Installation section above).

**File not found errors**

Run from the project root:

```bash
cd /path/to/NewAthenaE2EPSF
python3 gui_distributions.py
```

**"No MM selected" error**

Go to the MM Configuration tab, check at least one MM, and click **Apply Selection**.

**Parameters don't update after loading a preset**

In Standard mode, switch to the Generate sub-tab after loading. In Free mode, click **Generate PSF Data**.

**Export fails**

Close the Excel file if it is open in Excel. Check file permissions with `ls -la Distributions/`.

**Alpha values unexpected**

Values are automatically clamped to [0, 1]. Alpha controls are only visible for pseudo-voigt distribution type.

**Figures not saving**

```bash
mkdir -p Figures
```

**`Aeff_sum_mod` appears incorrect**

Run `.inspect_export.py` on the package. Check whether the `A_eff` sheet contains uncached formula cells. Use `tools/compute_aeff_values.py` to populate numeric values.

**Thermal perturbations all zero**

This can occur when the `d_therm_*` columns contain Excel formulas without cached values (e.g. `=U2`). The code includes a formula-inspection fallback that resolves the referenced column and uses its numeric data. If thermal data still appears zero, verify the referenced columns in the Thermal sheet contain numeric values.

---

## Contributing

- Run `pytest` before submitting changes
- Keep changes small and focused; preserve public function signatures where possible
- Report issues with a short reproduction: commands, stack trace, and a small Excel workbook

**Running tests:**

```bash
source .venv/bin/activate
python -m pytest -q
```

**Headless smoke test:**

```bash
python3 main.py -f Distributions/TestDistribution.xlsx -o /tmp/out.png
```

---

## Author

**Ivo Ferreira**
European Space Agency
ivo.ferreira@esa.int
ORCID: https://orcid.org/0000-0002-9501-862X

---

## Release History

### v9.2 (2026-05-22)

- **HEW Contribution Ranking panel** (bottom pane of MM Selector): vectorised leave-one-out analysis ranks all checked MMs by individual HEW contribution in typically < 1 s for the full 600-MM set. Results shown in a sortable table (Rank / MM # / delta-HEW / Row / Petal) coloured red (degrading) or green (improving). Multi-select with Shift-click and Cmd/Ctrl-click; right-click "Select in tree" syncs checkboxes to the highlighted rows.
- **HEW Contribution Map**: Map button opens a scatter plot of all MM positions colour-coded by delta-HEW (RdYlGn, symmetric around 0); unranked MMs shown as grey crosses; axis limits fixed to the full telescope footprint.
- Thermal formula-cell resolution fix: if `d_therm_*` columns contain uncached Excel formulas, the referenced TC column is resolved and used instead.
- 98 tests passing.

### v9.1 (2026-05-22)

- Vignetting `single` and `per_pos` table modes now correctly apply factors in both the row-by-row loop and the final reconcile pass (previously only `per_row_energy` was handled).
- All `np.interp()` calls in the vignetting path now use `abs(rotation_value)`; tables use non-negative delta values with symmetry assumed.
- 98 tests passing.

### v9 (2026-05-22)

- **Interactive MM Selector viewer**: running `main.py` without `--output` or `--export-package` opens a split window with a Row -> Petal -> MM checkbox tree and a live E2E PSF + EEF figure. Select All / None buttons and a partial-row indicator support bulk selection. `--export-package` bypasses the GUI entirely for headless runs.
- **Defocus PSF shape broadening**: per-MM `sigma_rad`/`sigma_azi` adjusted using a linear geometric model -- PSF size grows linearly with dz from best focus to the MM physical dimension at `dz = 12 m`.
- `plot_sum()` gains `return_fig` and `figsize` parameters.
- 98 tests passing (13 new).

### v8 (2026-04-17)

- **Off-axis pointing**: angle decomposed into X/Y rotations (`* 60 / sqrt(2)`) and written to the new `Extra PSF shifts` sheet.
- **Defocus**: written as `d_extra_z [um]` to `Extra PSF shifts`; projected to centroid shift via Z-axis formula.
- **HEW sigma broadening**: `MM HEW degradation rotazi/rotrad` sheets with per-position angle -> HEW interpolation and quadrature sigma broadening. Results written to `Extra PSF degradations` and MM_PSF I/J.
- **Energy-dependent sigma scaling**: `MM HEW degradation energy` sheet with energy/factor table applied after HEW broadening.
- **VLOOKUP resolver**: Python fallback for MM_PSF D/E formula cells; values persisted as plain numbers.
- **Batch combinations**: `--batch-combinations` CLI for headless multi-configuration runs with per-config ZIP packaging and aggregated results workbook.
- Preset table shifted from column K to **M** to free I/J columns.
- Pearson4 skipped in coarse mode.
- 71 tests passing (20 new).

### v7 (2026-04-14)

- Repository reorganisation: tests grouped under `tests/` by concern; utilities moved to `tools/`.
- Documentation and docstring pass across core modules.
- 42 tests passing.

### v6 (2026-04-01)

- GUI: improved combobox click behavior on macOS; MM Configuration checkbox toggling improved.
- Added `tools/compute_aeff_values.py` for caching A_eff numeric columns.
- Repo hygiene: removed generated artifacts from index; `.gitignore` updates.

### v5 (2026-02-07)

- GUI A_eff export writes numeric values to column B and clears column C.
- Percent-variable presets synthesized to explicit Gaussian forms.
- Preset energy tokens (e.g. `1 keV`) parsed and written to vignetting sheet `C2` at export.
- CLI preserves legacy column C behavior.
- Replaced debug prints with Python `logging` calls.

### v4 (2026-02-03)

- Fixed dz -> dm projection; correct per-MM dz outputs with `--log-dz`.
- Vignetting application re-ordered after A_eff initialisation with explicit bookkeeping (`aeff_base`, `aeff_adjusted`, `aeff_vig_factor`).
- Improved vignetting parsing for two-column and per-position formats.
- GUI: Apply vignetting factors when exporting checkbox in A_eff tab.
- 34 tests passing.

### v3 (2026-01-28)

- Repository cleanup and removal of legacy tooling.
- Deterministic per-MM sampling for presets; CSV/Excel export parity.
- Unit tests and pytest configuration added.
- Renamed `sensivitiy` to `sensitivity`.

### v2 (2025-12-15)

- Initial public release. Core PSF generation, placement strategies, Excel I/O, GUI and CLI analysis utilities, Gaussian and Pseudo-Voigt support.
