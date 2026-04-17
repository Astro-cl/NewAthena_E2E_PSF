# NewAthenaE2EPSF — PSF Analysis Toolkit

[![Python CI](https://github.com/Astro-cl/NewAthena_E2E_PSF/actions/workflows/python-ci.yml/badge.svg)](https://github.com/Astro-cl/NewAthena_E2E_PSF/actions/workflows/python-ci.yml)

A comprehensive toolkit for PSF (Point Spread Function) modeling and analysis of mirror module configurations with support for multiple distribution types, perturbation analysis, and an interactive GUI.

> **Documentation, comments and Unit tests written by AI.**

---

## Documentation Index

| Document | Description |
|----------|-------------|
| [README.md](README.md) | This file — main project overview and guide |
| [QUICKSTART.txt](QUICKSTART.txt) | Quick start instructions for common tasks |
| [DOCS_GUI.md](DOCS_GUI.md) | GUI user manual with screenshots and workflows |
| [DOCS_SUMMARY.md](DOCS_SUMMARY.md) | Summary of core modules and API reference |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Guidelines for contributors |
| [RELEASE_NOTES.md](RELEASE_NOTES.md) | Full release history and change log |
| [DOCS_FEATURES_APRIL2026.md](DOCS_FEATURES_APRIL2026.md) | v8 feature documentation (off-axis, defocus, HEW degradation, batch) |
| [SENSITIVITY_QUICKSTART.txt](SENSITIVITY_QUICKSTART.txt) | Sensitivity pipeline guide |

---

## Installation

### Requirements

- Python 3.8 or higher
- pip (Python package manager)

### Setup

1. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

   Required packages:
   - `numpy`
   - `pandas`
   - `matplotlib`
   - `openpyxl`

2. **Verify installation**:
   ```bash
   python3 --version   # Should be 3.8 or higher
   python3 -c "import numpy, pandas, matplotlib, openpyxl; print('All dependencies installed successfully')"
   ```

---

## Quick Start

The **core engine** is `main.py` (CLI). It reads a fully-configured Excel workbook and produces PSF plots, metrics, and export packages. The **GUI** (`gui_distributions.py`) is a helper tool whose primary purpose is to build and populate that Excel workbook — once the file is ready, the CLI drives all analysis.

### Using the Command Line

```bash
python3 main.py -f Distributions/your_file.xlsx
```

If `--output` is not passed, an interactive window opens with plot export shortcuts. If `--output` is passed, the figure is saved and the script exits without opening a window.

### Using the GUI — to prepare the input file

The GUI's main role is to configure and export the Excel workbook that `main.py` requires. Once the file is exported, hand it to the CLI for analysis.

1. **Launch the GUI**:
   ```bash
   python3 gui_distributions.py
   ```
   > **macOS trackpad note:** If you use tap-to-click, button activation is lenient (does not require a perfectly stationary press/release). For dropdowns (comboboxes), clicking anywhere on the field opens the list.

2. **Load your Excel file**:
   - Click **"Load Excel File"**
   - Navigate to the `Distributions/` folder
   - Select an Excel file with MM configuration

3. **Select Mirror Modules**:
   - In the **MM Configuration** tab, check/uncheck MMs to include
   - Use row/MM number filters to select specific subsets

4. **Generate PSF Data**:
   - Go to **PSF → Generate** tab
   - Choose distribution type (gaussian or pseudo-voigt)
   - Set parameters for each MM characteristic
   - Click **"Generate PSF Data"**

5. **Export Results**:
   - Go to the **Export** tab
   - Choose export mode (new file or update current)
   - Click **"Export to Excel"** — this produces the workbook consumed by `main.py`
   - Files are saved in `Distributions/`
   - Right-click on plots for: Export PSF / Encircled Energy / FITS / EEF CSV / Fit Parameters CSV

---

## Features

### Core Capabilities

- **Multiple Distribution Types**: Gaussian and Pseudo-Voigt (mixture of Gaussian and Lorentzian) with independent azimuthal and radial alpha parameters
- **Standard PSF Presets**: Load predefined MM_PSF distributions from the Excel preset table with automatic parameter derivation for variable presets
- **Row-Wise MM Optimizer**: Optimize MM# assignments within rows to minimize HEW while keeping physical locations fixed
- **Rotation-Invariant HEW**: Polar grid integration eliminates orientation bias
- **Perturbation Analysis**: Alignment errors, gravity offload effects, and thermal deformations
- **Off-Axis & Defocus Modeling**: X/Y off-axis decomposition and focus-shift projection to centroid offsets (v8)
- **HEW Degradation**: Per-position sigma broadening from lookup tables (Row #, angle, energy to HEW arcsec) (v8)
- **Batch Combinations**: Automated multi-configuration runs (off-axis, energy, defocus) with per-config ZIP packaging (v8)
- **Fast Mode**: Optimized computation with configurable sampling density
- **Multi-MM Analysis**: Process multiple mirror modules with different distributions
- **Excel Integration**: Load/save configurations with full parameter support and formula preservation
- **Interactive GUI**: User-friendly interface for parameter generation and configuration with standard/free preset modes
- **Figure Export**: High-resolution PNG exports of PSF and encircled energy plots

### Distribution Features

- **Fixed Values**: All mirror modules use identical parameters
- **Gaussian Distribution**: Parameters vary across MMs using normal distribution N(mu, sigma^2)
- **Uniform Distribution**: Parameters vary across MMs using uniform distribution U(min, max)
- **Alpha Parameter Control**: For Pseudo-Voigt, `alpha_rad` and `alpha_azi` can be fixed or vary per MM (automatically clamped to [0, 1])
- **Standard Presets**: Load predefined PSF distributions from the Excel preset table (MM_PSF sheet, starting at cell **M1**) with automatic variable parameter derivation
- **Free Mode**: Full manual control over all PSF parameters

---

## Usage

The typical workflow is: **GUI → export Excel → CLI**. The GUI builds the input workbook; the CLI reads it and performs all PSF analysis, metric calculation, and export packaging.

### Command-Line Interface (CLI)

```
python3 main.py [options]
```

| Option | Description |
|--------|-------------|
| `-f / --file` | Path to input Excel workbook |
| `--mode {coarse,fine}` | Computation mode (default: coarse) |
| `--output PATH` | Save combined figure to PATH and exit (no GUI window) |
| `--normalize` | Normalize PSF for plot comparison |
| `--export-package` | Package figures, FITS, and workbook into `Exports/<TIMESTAMP>/` |
| `--single-config N` | Only run configuration index N (useful for debugging) |
| `--optimize` | Enable row-wise MM# assignment optimization |
| `--batch-combinations` | Automated multi-configuration batch run (v8) |
| `--log-dz` | Log per-MM dz projections |

**Examples:**

```bash
# Quick coarse analysis, interactive window
python3 main.py -f Distributions/my_config.xlsx

# Save figure directly, no window
python3 main.py -f Distributions/my_config.xlsx --output Figures/result.png

# Coarse export package
python3 main.py --file Distributions/YourWorkbook.xlsx --export-package --mode coarse

# Fine-resolution export package (compute-intensive)
python3 main.py --file Distributions/YourWorkbook.xlsx --export-package --mode fine

# Debug: run only config index 1
python3 main.py --file Distributions/YourWorkbook.xlsx --export-package --single-config 1

# Batch combinations (off-axis, energy, defocus permutations)
python3 main.py --file Distributions/YourWorkbook.xlsx --batch-combinations --mode coarse

# Optimize MM# row assignments
python3 main.py -f Distributions/my_config.xlsx --optimize
```

**Background / unattended runs:**

```bash
nohup python3 main.py --file Distributions/YourWorkbook.xlsx --export-package --mode fine &
# or with screen:
screen -S e2e_run
python3 main.py --file Distributions/YourWorkbook.xlsx --export-package --mode fine
# detach with Ctrl-A D
```

---

### GUI Application — building the input workbook

The GUI's primary purpose is to configure mirror-module PSF parameters and export them into the Excel workbook that `main.py` expects. It is **not** the analysis engine — once the workbook is ready, run the CLI. The GUI also provides a live PSF preview and right-click export shortcuts.

Launch with:

```bash
python3 gui_distributions.py
```

**Main Window Layout:**
```
+------------------------------------------------------------------+
|  New Athena E2E - PSF Configuration Tool              [_][O][x] |
+------------------------------------------------------------------+
|  [Load Excel File]  Current: None                                |
+------------------------------------------------------------------+
|  +------------------------------------------------------------+  |
|  | [MM Configuration] [PSF] [Other Data Types] [Export]      |  |
|  |                                                            |  |
|  |                  (Tab content area)                        |  |
|  |                                                            |  |
|  +------------------------------------------------------------+  |
+------------------------------------------------------------------+
```

#### MM Configuration Tab

Shows all MMs from your Excel file in a scrollable table.

**Features:**
- **Checkboxes**: Select which MMs to include in PSF generation
- **Select/Deselect All**: Toggle all checkboxes at once
- **Filters**: Show only specific rows or MM numbers; filter combinations (Row# = 1, MM# = All gives all MMs in row 1)
- **Sortable columns**: Click column headers to sort
- **Apply Selection**: Must click to confirm selection changes

#### PSF Tab — Standard Mode

Use standard presets for quick, consistent PSF generation.

**Sub-tabs:** `[Select]` `[Generate]`

**Steps:**
1. Ensure **"Standard"** radio button is selected (default)
2. Click dropdown labelled **"Load from Excel (M1):"** to see available presets
3. Select desired preset
4. Click **"Load Standard Distribution"**
5. Switch to the **Generate** sub-tab to see auto-filled parameters
6. Click **"Generate PSF Data"**

**Preset Types:**
- **Fixed presets** (no %): All MMs get identical values
- **Variable presets** (with %): Parameters vary per MM using Gaussian distribution; sigma_mean derived from HEW, sigma_sigma = sigma_mean × percent
- **Pseudo-Voigt presets** (with ratio): Include `alpha_rad`/`alpha_azi` values

In Standard mode, parameters are read-only (grayed out) in the Generate tab. Click **"Generate PSF Data"** to apply to selected MMs.

#### PSF Tab — Free Mode

Full manual control over all parameters.

1. In **PSF → Select**, choose **"Free"** radio button
2. In **PSF → Generate**, select distribution type (gaussian / pseudo-voigt)
3. For each parameter, choose distribution (fixed / gaussian / uniform) and set values:
   - **Fixed**: 1 field (value)
   - **Gaussian**: 2 fields (mean, sigma)
   - **Uniform**: 2 fields (min, max)
4. Alpha parameters only appear for pseudo-voigt; values are auto-clamped to [0, 1]
5. Click **"Generate PSF Data"**

**Example configurations:**

| Goal | Settings |
|------|---------|
| All MMs identical | All parameters: fixed |
| Varying widths only | m_rad, m_azi: fixed=0; sigma_*: gaussian |
| Asymmetric pseudo-voigt | Type: pseudo-voigt; alpha_rad: fixed=0.8; alpha_azi: fixed=0.2 |

#### Other Data Types Tab

Generate additional perturbation data alongside the PSF.

**Available Types:**
- **MM_alignement**: Manufacturing/assembly errors (dz, dx, dy, rotations)
- **MM_gravity**: Gravity-induced deformations
- **MM_thermal**: Thermal expansion/contraction

Each type uses the same fixed/gaussian/uniform parameter scheme. Click **"Generate Selected Data"** for each type needed. All generated data accumulates for a single export operation.

#### Export Tab

**Export Modes:**
- **Update Current File**: Overwrites generated sheets; preserves formulas, formatting, and other sheets; creates a backup before writing.
- **Create New File**: Saves to a new filename; original file unchanged; auto-suggests name `original_name_new.xlsx`.

**A_eff export behavior (GUI vs CLI):**
- GUI export evaluates standard A_eff presets per-MM and writes numeric values into column B of the `A_eff` sheet; column C is cleared.
- CLI (`main.py`) preserves legacy behavior: adjusted/derived A_eff is written to column C.

---

### Keyboard Shortcuts & Plot Context Menu

**GUI navigation:**
- `Tab` — navigate between fields
- `Enter` — activate focused button
- `Esc` — close dialogs
- Click column headers to sort tables
- Scroll with mouse wheel or trackpad

**Plot context menu (right-click on plot):**
- Export PSF PNG (high-resolution)
- Export Encircled Energy (EEF) PNG
- Export EEF CSV (saved to `CustomPSFs/E2E_EEF_YYYYMMDD_HHMMSS.csv`)
- Export Fit Parameters CSV
- Export FITS (PSF matrix)

---

### Row-Wise MM Optimizer

The optimizer finds the best MM# assignment within each physical row of the mirror while keeping MM positions fixed, minimizing the aggregate HEW.

**Placement strategies** (controls how ranked MMs are placed in angular slots):

- `random` — random assignment (baseline)
- `best_center` — places the best MMs closest to the optical axis center
- `worst_center` — places the worst MMs at center; best MMs at edges
- `alternating` — alternates best/worst from center outward
- `cross` — places the best MMs on a 90-degree cross-like pattern (+-x/+-y) to reduce anisotropy
- `x_axis` — places best MMs preferentially near the +-x direction, alternating above/below to avoid clustering
- `elliptical` — within each row assigns best MMs to slots closer to the x-axis and worst to slots closer to y-axis

**Usage:**

```bash
python3 main.py -f Distributions/input.xlsx --mode coarse --optimize
```

Output files:
- Optimized workbook: `Distributions/my_data_optimised.xlsx`
- Plots: `Figures/E2E_PSF_*_optimised_*.png`

**Example workflow:**

```bash
# Baseline analysis
python3 main.py -f Distributions/my_config.xlsx
# HEW: 15.2 arcsec (baseline)

# Optimize MM# assignments
python3 main.py -f Distributions/my_config.xlsx --optimize
# HEW: 14.1 arcsec (optimized)
```

When to use: when different MM# have different PSF characteristics, when MMs within a row must stay in that row but can swap positions, and when you want to minimize HEW without redesigning the system.

---

## Distribution Types

### Gaussian Distribution

Standard 2D Gaussian function:

```
G(x,y) = A * exp(-0.5 * r^2)
```

where `r^2` is the rotated squared distance from the center.

**Characteristics:**
- Sharp central peak; exponential decay in tails
- Suitable for well-aligned, stable systems

### Pseudo-Voigt Distribution

Product of two independent 1D Pseudo-Voigt functions (azimuthal x radial):

```
PV(azi, rad) = PV_azi(azi) x PV_rad(rad)
PV_1D(x) = (1-alpha) * G(x) + alpha * L(x)
```

where:
- `alpha` is the mixing parameter in [0, 1]
- `alpha = 0`: Pure Gaussian (sharp peak)
- `alpha = 0.5`: Balanced mix
- `alpha = 1`: Pure Lorentzian (heavy tails)
- Independent `alpha_rad` and `alpha_azi` for asymmetric PSFs

**Alpha values guide:**

```
alpha = 0.0  Pure Gaussian (sharp core)
alpha = 0.2  Mostly Gaussian
alpha = 0.5  Balanced (50/50 mix)
alpha = 0.7  More tails (scattering effects)
alpha = 1.0  Pure Lorentzian (heavy tails)
```

**Use cases:**
- **Low alpha (0.0 to 0.3)**: Good alignment, minimal scattering
- **Medium alpha (0.3 to 0.7)**: Moderate scattering effects
- **High alpha (0.7 to 1.0)**: Significant scattering, degraded optics
- **Asymmetric (alpha_rad != alpha_azi)**: Directional scattering or optical aberrations

### Modified Pseudo-Voigt (Aggregated Radial Fit)

The end-to-end fit in `main.py` applies a modified pseudo-Voigt model to the azimuthal-average radial intensity profile of the aggregated PSF. This combines a narrow Gaussian core with a broader wing-shaped term:

```
G(r; Gamma_c) = exp(-4 ln2 (r / Gamma_c)^2)
C(r; Gamma_w) = [1 + a (2r / Gamma_w)^2]^{-beta}    where a = 2^(1/beta) - 1
mix  = (1 - eta) G(r; Gamma_c) + eta * scalar * C(r; Gamma_w)
norm = (1 - eta) + eta * scalar
I(r) = A * mix / norm
```

**Parameters:**

| Symbol | Name | Description |
|--------|------|-------------|
| `A` | Amplitude | Overall intensity scale of the radial profile |
| `Gamma_c` | `Gamma_core` | Core width (arcsec); controls the narrow Gaussian-like peak |
| `Gamma_w` | `Gamma_wing` | Wing width (arcsec); controls the broader tail scale |
| `eta` | eta | Mixing fraction; `eta = 0` gives pure core, `eta = 1` gives wing-dominated |
| `beta` | beta | Wing shape exponent; `beta = 1` gives Lorentzian-like, larger values give faster-decaying |
| `scalar` | — | Additional scaling of the wing component before normalization |

Normalization by `(1 - eta) + eta * scalar` preserves the overall amplitude. The fit result is saved as `Figures/E2E_fit.png`.

---

## Physical Model & Formulas

### Rotated 2D Gaussian

Let `(x, y)` be evaluation coordinates and `(mu_x, mu_y)` the center, with principal standard deviations `sigma_x`, `sigma_y` and rotation angle `theta` (radians):

```
dx = x - mu_x;    dy = y - mu_y
c = cos(theta);   s = sin(theta)

a     = c^2/sigma_x^2 + s^2/sigma_y^2
b     = s*c*(1/sigma_x^2 - 1/sigma_y^2)
ccoef = s^2/sigma_x^2 + c^2/sigma_y^2

E      = -0.5 * (a*dx^2 + 2*b*dx*dy + ccoef*dy^2)
Output = A * coeff * exp(E)
```

where `coeff = 1/(2*pi*sigma_x*sigma_y)` if `normalize=True`, else `coeff = 1`.

**Parameters:** `mux`, `muy` (center, meters); `sigmax`, `sigmay` (principal sigmas, meters); `theta` (rotation, radians); `amplitude`; `normalize`.

Rotation conventions are consistent across `gaussian_2d_rotated`, `pseudo_voigt_2d_rotated`, and `eval_psf_matrix_rotated` so analytic and discrete PSF evaluations align.

Public APIs use **meters**; the loader converts arcsec using `1 arcsec = 12 * pi / 180 / 3600` meters.

### Separable Pseudo-Voigt

Each axis uses a 1D Pseudo-Voigt:

```
PV(u; sigma, alpha) = (1 - alpha) * G(u; sigma) + alpha * L(u; sigma)
```

where `G(u; sigma) = (1/sqrt(2*pi)) * exp(-u^2/2)` and `L(u; sigma) = (1/pi) * 1/(1 + u^2)`.

The 2D shape is `PV_azi(azi_rot/sigma_azi) * PV_rad(rad_rot/sigma_rad)` evaluated in the rotated principal-axis frame.

- `alpha` may be specified per-axis (`alpha_azi`, `alpha_rad`) or a single `eta` mixing parameter may be used as fallback.
- `normalize=True` divides by `(sigma_azi * sigma_rad)` to account for the change of variables.

### Pearson Type IV

The 1D unnormalized profile is:

```
P(u) proportional to (1 + (u/sigma)^2)^{-m} * exp(nu * atan(u/sigma)),    u = x - mu
```

**Parameters:** `mu` (center, meters); `sigma` (scale, meters); `m` (tail-shape, > 0); `nu` (skew/asymmetry; `nu = 0` gives symmetric).

In 2D the implementation uses separable azimuthal/radial forms or a radial-only fit depending on the target data; normalization constants are computed numerically when required.

> Note: Pearson 4 is skipped in coarse/quick mode for performance.

### King / Moffat-like Profile

Radial form:

```
K(r) = A * (1 + (r/alpha)^2)^{-beta}
```

**Parameters:** `A` (amplitude); `alpha` (core scale, meters); `beta` (wing exponent, > 1).

Supports rotationally symmetric 2D evaluation; normalization computed analytically or numerically as appropriate.

---

### Fitting Pipeline

The end-to-end fitting routine in `main.py` runs automatically after PSF aggregation and proceeds in the stages below.

#### Stage 1 — PSF Aggregation

Individual mirror-module (MM) PSFs are co-added on a shared 2D Cartesian grid:

```
Z(x, y) = sum_i  w_i * PSF_i(x - mu_xi, y - mu_yi)
```

where `w_i` is the optional weight and `(mu_xi, mu_yi)` is the centroid of module `i`. All coordinates are in meters; the chosen distribution type (Gaussian, Pseudo-Voigt, …) determines the shape of each `PSF_i`.

#### Stage 2 — Best-focus Minimisation

The optimal focal-plane centre `(cx, cy)` minimising the HEW is located by a discrete gradient search starting from the nominal centre, using 1 μm steps and up to 30 iterations (or 20 in fast mode). The search is seeded from multiple candidate starts.

#### Stage 3 — Rotation-invariant Radial Profile

The azimuthally averaged radial intensity profile `I(r)` is computed by polar-grid sampling centred at `(cx, cy)`:

```
E(r_k) = sum_{j} Z(cx + r_k cos theta_j, cy + r_k sin theta_j) * r_k * Delta_theta

Phi(r)  = cumsum_k  E(r_k) * Delta_r          (cumulative radial energy)

I(r)    = E(r) / (2 * pi * r)                  (mean radial intensity)
```

Default grid: `N_r = 400`, `N_theta = 360`, `r_max = r_centres + 5 * max(sigma)`.

The profile is adaptive: if less than 99.95 % of the total energy is enclosed in `r_max`, the grid expands up to 3 times (factor 1.5 per iteration) to capture heavy tails.

#### Stage 4 — Encircled Energy Function (EEF) and HEW

The normalised Encircled Energy Function is:

```
EEF(d) = Phi(d/2) / Phi(r_max)
```

Radii enclosing 50%, 80%, and 90% of the energy are located by local cubic interpolation of the EEF curve. The corresponding diameters are:

```
HEW    = 2 * r_50%     (Half-Energy Width)
EEF-80 = 2 * r_80%
EEF-90 = 2 * r_90%
```

All diameters are converted to arcsec using `1 arcsec = 12 * pi / 180 / 3600 m`.

#### Stage 5 — Modified Pseudo-Voigt Radial Fit

The intensity profile `I(r)` is fitted in two stages.

**Stage 5a — Core Gaussian (seed estimate):** A simple Gaussian is fitted to the core region (`r <= 2 * Gamma_0`) to give robust initial values for amplitude `A` and core width `Gamma_c`:

```
I_G(r) = A * exp(-4 ln 2 * (r / Gamma_c)^2) + b
```

**Stage 5b — Full core + wing fit:** The modified pseudo-Voigt model is fitted to the full profile:

```
G(r; Gamma_c)  = exp(-4 ln 2 * (r / Gamma_c)^2)

a              = 2^(1/beta) - 1
C(r; Gamma_w)  = [1 + a * (2r / Gamma_w)^2]^{-beta}

mix            = (1 - eta) * G(r; Gamma_c) + eta * scalar * C(r; Gamma_w)
norm           = (1 - eta) + eta * scalar

I(r)           = A * mix / norm
```

The **objective** minimises EEF residuals (not intensity directly):

```
EEF_model(r) = cumsum(2 pi r I_model(r) dr) / total_energy_model

residual(r)  = EEF_model(r) - EEF_data(r)
```

The optimisation uses `scipy.optimize.least_squares` with `soft_l1` robust loss and **36 multi-start** perturbations (±12 % random offsets about the seed). Parameter bounds:

| Parameter | Lower | Upper |
|-----------|-------|-------|
| `A` | 0.5 × A₀ | ∞ |
| `Gamma_c` | 0.2 × Γ_c,0 | 3 × Γ_c,0 |
| `Gamma_w` | 1.2 × Γ_c,0 | 25 × Γ_c,0 |
| `eta` | 0.05 | 0.50 |
| `beta` | 1.0 | 5.0 |
| `scalar` | 0.2 | 12.0 |

Fit results (A, Gamma_core, Gamma_wing, eta, beta, scalar) are exported to CSV via `export_fit_params_csv()`.

#### Stage 6 — Pearson Type IV Fit

A Pearson Type IV model (`lmfit.models.Pearson4Model`) is fitted using the same EEF-objective framework:

```
P(x; mu, sigma, m, nu)  proportional to
    (1 + ((x - mu)/sigma)^2)^{-m}  *  exp(-nu * atan((x - mu)/sigma))
```

| Symbol | Parameter | Description |
|--------|-----------|-------------|
| `mu` | center | Profile centre (arcsec) |
| `sigma` | sigma | Scale width (arcsec) |
| `m` | expon | Tail exponent (m > 1; larger = faster decay) |
| `nu` | skew | Asymmetry; `nu = 0` gives a symmetric profile |

Bounds: `m ∈ [1, 8]`, `nu ∈ [-2, 2]`. The same 36-start EEF least-squares strategy is applied. A quick intensity-only `lmfit` fit seeds the multi-start. Pearson IV is **skipped in quick/coarse mode**.

#### Stage 7 — FWHM from 1D Marginals

FWHM is extracted from the 1D marginal profiles of the aggregated 2D PSF:

```
prof_x(x) = integral_y  Z(x, y) dy
prof_y(y) = integral_x  Z(x, y) dx

FWHM_x = x_{half-max, right} - x_{half-max, left}
FWHM_y = y_{half-max, right} - y_{half-max, left}
```

#### Stage 8 — Directional HEW from Marginals

`HEW_x` and `HEW_y` are the minimum-width intervals of each 1D marginal containing exactly 50 % of the total marginal energy:

```
HEW_x = min{ |b - a| : integral_a^b prof_x(x) dx >= 0.5 * integral prof_x dx }
```

and equivalently for `HEW_y`.

#### Output Metrics Summary

| Metric | Description |
|--------|-------------|
| `HEW` | Half-Energy Width diameter (arcsec), rotation-invariant |
| `HEW @(0,0)` | HEW at the focal-plane origin |
| `EEF-80` | Diameter enclosing 80 % of energy (arcsec) |
| `EEF-90` | Diameter enclosing 90 % of energy (arcsec) |
| `FWHM_x` / `FWHM_y` | Half-maximum width of each marginal (arcsec) |
| `HEW_x` / `HEW_y` | Marginal half-energy widths (arcsec) |
| `Gamma_core` | Modified PV core FWHM parameter, arcsec |
| `Gamma_wing` | Modified PV wing width parameter, arcsec |
| `eta` | PV core-to-wing mixing fraction |
| `beta` | PV wing shape exponent |
| `scalar` | PV wing amplitude scale |

---

### Perturbation Model

#### d_z to Centroid Shift Projection

An axial displacement `d_z` of a mirror module produces a centroid shift in the focal plane via:

```
Delta = d_z * x_MM / (f - z_MM)
```

where `f = 12 m` is the focal length and `x_MM`, `z_MM` are the mirror module coordinates.
This projection is applied to alignment, gravity, thermal, and defocus perturbations.

#### Off-Axis Decomposition (v8)

Off-axis pointing is decomposed into equal X and Y rotation components:

```
off_x = off_axis_deg * 60 / sqrt(2)   [arcsec]
off_y = off_axis_deg * 60 / sqrt(2)   [arcsec]
```

These are written to a dedicated **"Extra PSF shifts"** sheet and applied additively in `compute_total_rot_polar()`.

Defocus is written as `d_extra_z [um]` to the same sheet (mm * 1e3 gives um), then projected to centroid shifts via the d_z formula above.

#### HEW Sigma Broadening (v8)

Per-position HEW degradation is computed from lookup tables (Row #, angle, energy gives HEW arcsec) by interpolation. The broadened sigma is:

```
sigma_new = sqrt(sigma_base^2 + (HEW / (2*sqrt(2*ln(2))))^2)
```

Results are written to the **"Extra PSF degradations"** sheet (`sigma_extra`) and MM_PSF columns I/J (degraded sigma).

### Unit Conventions

- **Arcsec to meters:** `1 arcsec = 12 * pi / 180 / 3600 m`  
  (chosen to match the NewAthena optical geometry; focal length `f = 12 m`)
- **HEW to sigma (symmetric Gaussian):** `sigma = HEW / (2*sqrt(2*ln(2))) ~ HEW / 2.3548`
- **Asymmetric Gaussian:** For an elliptical 2D Gaussian the 50% encircled-energy diameter depends on both `sigma_rad` and `sigma_azi`; there is no closed-form `sigma = HEW/2.355` mapping. Preset table values should already encode the encircled-energy sigma.

**Reference sigma values (arcsec) for common presets:**

| Preset | sigma_rad | sigma_azi |
|--------|-----------|-----------|
| Symmetric Gaussian, 4.3 arcsec HEW | 1.826058874 | 1.826058874 |
| Symmetric Gaussian, 8 arcsec HEW | 3.397323952 | 3.397323952 |
| Asymmetric Gaussian, 4.3 arcsec HEW, ratio 4:1 | 2.963740901 | 0.740935225 |
| Asymmetric Gaussian, 8 arcsec HEW, ratio 7:1 | 5.797205089 | 0.828172156 |

---

## User Manual

### Standard MM_PSF Presets

The GUI loads predefined distributions from the Excel preset table in the MM_PSF sheet, starting at cell **M1**. This is the source of truth for the "Standard" preset dropdown.

In **Standard mode** the GUI:
- Reads the preset table from the loaded Excel file
- Populates `sigma_rad` / `sigma_azi` (and pseudo-voigt `alpha_*`) controls from that row
- Forces `m_rad` and `m_azi` to **fixed 0** (the standard table does not define them)

In **Free mode**, all parameters are under full manual control.

#### Preset Table Structure (MM_PSF sheet, cell M1)

Columns (left to right):
1. Preset Name (text)
2. `sigma_rad` spec
3. `sigma_azi` spec
4. `alpha_rad` spec (pseudo-voigt only; use `-` or blank otherwise)
5. `alpha_azi` spec (pseudo-voigt only; use `-` or blank otherwise)

Each spec cell can be:
- A plain number — interpreted as **fixed**, e.g. `3.397323952`
- A Gaussian spec: `gaussian(mean, sigma)`, e.g. `gaussian(3.397, 10%*3.397)`
- A Uniform spec: `uniform(min, max)`, e.g. `uniform(3.0, 4.0)`

> All values are in **arcsec**.

#### Common Pitfalls

- **You must reload the workbook** in the GUI after editing the preset table in Excel.
- **Don't leave spec cells blank**: blank `sigma_rad`/`sigma_azi` may trigger fallback behavior.
- **Use dot decimals**: write `3.397`, not `3,397`.
- **Alpha columns for Gaussians**: use `-` or leave blank; alpha values are only applied for pseudo-voigt.
- **Asymmetric Gaussian HEW is not sigma = HEW/2.355**: provide sigma values that match the 50% encircled-energy diameter (see reference table above).
- **If a pseudo-voigt preset has no alpha**: check that the preset name includes "pseudo-voigt" or "voigt" so it is classified correctly.

#### Adding Custom Presets

1. Open your Excel file in `Distributions/`
2. Locate the preset table (MM_PSF sheet, starting at cell M1)
3. Add/edit rows with the desired numeric/spec values
4. Save the Excel file
5. In the GUI, reload that file (the dropdown is refreshed on load)

---

### Batch Combinations (CLI, v8)

The `--batch-combinations` flag runs automated multi-configuration PSF analysis, sweeping over off-axis angles, X-ray energies, and defocus values.

```bash
python3 main.py --file Distributions/YourWorkbook.xlsx --batch-combinations --mode coarse
```

**What it does:**
- Generates all permutations of off-axis, energy, and defocus parameters defined in the workbook
- For each combination, writes the relevant perturbations to the workbook
- Runs a full PSF analysis per combination (headless, no GUI blocking)
- Packages each result into a ZIP under `Exports/`
- Writes an aggregated results workbook at the end

For full details see [DOCS_FEATURES_APRIL2026.md](DOCS_FEATURES_APRIL2026.md).

---

### Export Packages

Packages are written to `Exports/<TIMESTAMP>/` and include:
- Packaged workbook (authoritative source for `Aeff_sum_orig` / `Aeff_sum_mod`)
- Figures in `Figures/`
- FITS files in `CustomPSFs/`
- Aggregated Excel `E2E_EEF_and_fitparams_*.xlsx`

**Verification after package export:**
- PNGs in `Figures/` (coarse = 320x320 px, fine = 2062x2062 px)
- FITS in `CustomPSFs/` (grid dimensions = requested pixel size)
- Aggregated Excel present in the package

The CLI prefers the workbook copied into the package when computing `Aeff_sum_orig` and `Aeff_sum_mod`.

---

### A_eff Handling

- GUI export evaluates standard A_eff presets per-MM and writes numeric values to column B of the `A_eff` sheet; column C is cleared.
- CLI preserves legacy behavior: adjusted/derived A_eff is recorded in column C.
- When a preset name contains an energy token (e.g. `1 keV`), the GUI parses that energy and writes it into cell `C2` of the vignetting sheets on export.
- Vignetting application can be enabled via the **"Apply vignetting factors when exporting"** checkbox in the A_eff tab.
- **Percent-variable presets** (e.g. `Variable 10% 1 keV`) are synthesized to explicit gaussian forms internally so the evaluator can sample correctly.

#### Vignetting

- Two-column format (delta to factor) and per-position column formats are both supported.
- Vignetting application runs after A_eff initialization with explicit bookkeeping: `aeff_base`, `aeff_adjusted`, and `aeff_vig_factor`.
- MM300 receives the combined vignette factor when multiple factors apply (e.g. 0.1 x 0.1 = 0.01).

---

### Formula Evaluation

Many input workbooks use Excel formulas (VLOOKUP/XLOOKUP or textual preset expressions). When Excel cached numeric values are missing, an internal Python evaluator resolves common patterns.

**VLOOKUP resolver (v8):** A Python-based fallback resolves MM_PSF D/E formula values when openpyxl strips cached values; also persists base sigma as plain numbers for round-trip stability.

**Precomputing A_eff values:**

```bash
python3 tools/compute_aeff_values.py
```

This writes numeric A_eff into the `A_eff` sheet so downstream runs do not depend on formula evaluation.

---

### Sensitivity Pipeline

The repository includes a sensitivity-run template and driver under `sensitivity/`.

- **Per-combo input workbooks** are written to `sensitivity/input/` as `TIMESTAMP_index_<combo>.xlsx`. The runner prunes that folder to keep only the newest 100 files.
- **MM_PSF edits:** when writing/expanding the `MM_PSF` sheet, only per-MM input columns `B..H` are modified. Columns to the right (template/tail columns) are preserved.
- **Alpha masking:** for non-pseudo-voigt presets (Gaussian/Uniform), `alpha_rad` and `alpha_azi` are set to `-` (columns G and H).
- **Alignment / Thermal / Gravity zeroing:** when a combo requests `Alignment=0`, the runner zeros `d_align_rotazi` and `d_align_rotrad` in the `Alignment` sheet (columns B..G only). Similarly for `Thermal=0` and `Gravity offload=0`.
- **Partial results:** a per-job summary row is appended to `sensitivity/results/sensitivity_run_partial.csv` as jobs complete. Final consolidated results are written to `sensitivity/results/sensitivity_run_results.xlsx`.

**Run modes:**

```bash
# Generate input workbooks only (no execution)
python3 sensitivity/sensitivity_run.py --generate-only --baseline <file>

# Full run (generate inputs + execute all jobs)
python3 sensitivity/sensitivity_run.py --baseline <file>
```

For the full guide see [SENSITIVITY_QUICKSTART.txt](SENSITIVITY_QUICKSTART.txt).

---

### Repairing Existing Packages

If you have packages created before recent fixes, use the repair/diagnostic scripts at the repository root:

| Script | Purpose |
|--------|---------|
| `.inspect_export.py` | Inspect a package and list mismatches between packaged workbook and recorded aggregated sums |
| `.diagnose_aeff.py` | Identify formula-only A_eff columns and missing cached values |
| `.patch_fitparams.py` | Patch per-config `fitparams_aeffloss.xlsx` files inside packages |
| `.patch_aggregated.py` | Repair aggregated workbook rows that used the wrong A_eff source |

**Example:**

```bash
python3 .inspect_export.py Exports/20260416_124558
python3 .patch_fitparams.py Exports/20260416_124558 --dry-run
python3 .patch_fitparams.py Exports/20260416_124558
```

---

### File Formats

#### Input Excel File Structure

Required sheets:

**MM configuration:**
```
| MM # | Row # | x_MM [m] | y_MM [m] | z_MM [m] | r_MM [m] |
|------|-------|----------|----------|----------|----------|
|  1   |   1   |  0.123   |  0.456   |  7.89    |  0.15    |
```

**MM_PSF** (generated by GUI):
```
| MM # | m_rad ["] | m_azi ["] | sigma_rad ["] | sigma_azi ["] | distribution | alpha_rad | alpha_azi |
|------|-----------|-----------|---------------|---------------|--------------|-----------|----------|
|  1   |     0     |    5      |     8.2       |     7.8       | pseudo-voigt |   0.77    |   0.29   |
```

**Alignment** (optional):
```
| Position # | d_align_rad [um] | d_align_azi [um] | d_align_z [um] | d_align_rotz [arcsec] |
```

**Gravity offload / Thermal** (optional): same column pattern with `d_grav_*` / `d_therm_*` prefixes.

**Axis interpretation:**
- `x`, `y` — lateral displacements in x and y
- `z` — axial / focus displacement
- `rotz` — rotation about the optical axis
- `rad` — radial direction (positive = outward / higher radius)
- `azi` — azimuthal direction (positive = clockwise relative to the radial vector)

**Sign conventions:** positive `rotz` rotations introduce positive azimuthal shifts.

#### Output Files

| File | Description |
|------|-------------|
| `Figures/E2E_PSF_YYYYMMDD_HHMMSS.png` | PSF plot at 300 DPI |
| `Figures/Encircled_Energy_YYYYMMDD_HHMMSS.png` | EEF plot at 300 DPI |
| `Figures/E2E_fit.png` | Aggregated radial fit |
| `CustomPSFs/E2E_EEF_YYYYMMDD_HHMMSS.csv` | EEF CSV |
| `CustomPSFs/E2E_aggregated_*.fits` | PSF matrix in FITS format |

---

### Directory Structure

```
NewAthenaE2EPSF/
├── README.md                      # This file
├── requirements.txt               # Python dependencies
├── main.py                        # CLI analyzer and plotter
├── gui_distributions.py           # Interactive GUI application
├── distributions_rotated.py       # Core distribution functions
├── optimize_mm_rows.py            # MM optimizer + placement + Excel writer
├── Distributions/                 # Excel spreadsheet location
│   └── TestDistribution.xlsx      # Example file
├── Figures/                       # Exported plots location
├── CustomPSFs/                    # FITS and EEF CSV outputs
├── Exports/                       # Export packages (generated)
├── sensitivity/                   # Sensitivity pipeline driver and templates
├── tools/                         # Helper and diagnostic scripts
└── tests/                         # Unit and integration tests
```

**Core modules:**
- `main.py` — CLI entrypoints and Excel I/O helpers. Implements robust readers for `MM_PSF` and `A_eff` tables, spreadsheet-preserving export paths, PSF parameter conversion (arcsec to meters), vignetting application, and headless plot/export options used in CI and batch runs.
- `gui_distributions.py` — Tkinter GUI for interactive generation and exporting of per-MM distributions. Loads standard presets from the workbook, allows per-data-type distribution editing, previews sampled tables, and performs export with optional vignetting copy.
- `distributions_rotated.py` — Core distribution functions (Gaussian, Pseudo-Voigt, Pearson4, King).
- `optimize_mm_rows.py` — MM row optimizer, placement strategies, and Excel write helpers.

---

### Troubleshooting

**GUI won't start:**
```bash
python3 --version            # Must be 3.8+
pip install -r requirements.txt
```

**File not found errors:**
```bash
cd "/path/to/NewAthenaE2EPSF"
python3 gui_distributions.py
```

**"No MM selected" error:**
- Go to MM Configuration tab, check at least one MM, click "Apply Selection"

**Parameters don't update after loading a preset:**
- In Standard mode, switch to the Generate sub-tab after loading the preset
- In Free mode, ensure you clicked "Generate PSF Data"

**Export fails:**
- Close the Excel file if it is open in Excel
- Check file permissions; verify path with `ls -la Distributions/`

**Alpha values unexpected:**
- Values are automatically clamped to [0, 1]
- Alpha controls are hidden for gaussian distribution type (visible only for pseudo-voigt)

**Figures not saving:**
```bash
ls -la Figures/
mkdir -p Figures
```

**Exported package `Aeff_sum_mod` appears incorrect:**
- Re-run `.inspect_export.py` on the package
- Check whether the `A_eff` sheet contains formulas without cached values
- Use `tools/compute_aeff_values.py` or re-export from Excel to populate cached values

---

## Author

- **Ivo Ferreira** — primary author and maintainer
- **Affiliation:** European Space Agency
- **Contact:** ivo.ferreira@esa.int
- **ORCID:** https://orcid.org/0000-0002-9501-862X

---

## Release History

Full release notes are in [RELEASE_NOTES.md](RELEASE_NOTES.md).

### v8 (2026-04-17)
Major feature release. Off-axis pointing decomposed into X/Y components; defocus projected to centroid shifts; HEW sigma broadening from per-position lookup tables; Python VLOOKUP resolver for MM_PSF D/E formula values; `--batch-combinations` CLI for automated multi-configuration runs with per-config ZIP packaging and aggregated results workbook; improved A_eff formula evaluation fallback; Pearson4 skipped in coarse mode; preset table shifted from column K to **M** to avoid conflict with new I/J degraded sigma columns. 71 tests pass (20 new integration tests).

### v7 (2026-04-14)
Repository reorganization and test refactor. Moved utilities to `tools/`, grouped tests by concern under `tests/` subfolders, updated module docstrings. Local test run: 42 passed, 4 warnings.

### v6 (2026-04-01)
GUI polish (macOS combobox click behavior, MM Configuration checkbox reliability). Added `tools/compute_aeff_values.py`. Repo hygiene: removed generated files from repository index, added `.gitignore` entries.

### v5 (2026-02-07)
GUI A_eff export writes numeric values to column B and clears column C. Percent-variable presets synthesized to gaussian forms. Energy token parsing writes numeric energy to vignetting `C2`. CLI preserves legacy column C behavior.

### v4 (2026-02-03)
Fixed d_z to d_m projection. Re-ordered vignetting application with explicit bookkeeping (`aeff_base`, `aeff_adjusted`, `aeff_vig_factor`). Improved vignetting parsing for two-column and per-position formats. Extended interactive plot context menu with fit parameters CSV and EEF overlay. 34 tests pass.

### v3 (2026-01-28)
Repository cleanup. Deterministic per-MM sampling for presets. CSV/Excel parity in generation. Unit tests and pytest configuration added. Renamed `sensivitiy` to `sensitivity`.

### v2 (2025-12-15)
Initial public release. Core PSF generation, placement strategies, Excel I/O. Gaussian and Pseudo-Voigt support. Basic GUI and command-line analysis utilities.
