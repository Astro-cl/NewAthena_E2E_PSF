# NewAthenaE2EPSF_v3 - PSF Analysis Toolkit

[![Python CI](https://github.com/Astro-cl/NewAthena_E2E_PSF/actions/workflows/python-ci.yml/badge.svg)](https://github.com/Astro-cl/NewAthena_E2E_PSF/actions/workflows/python-ci.yml)

A comprehensive toolkit for PSF (Point Spread Function) modeling and analysis of mirror module configurations with support for multiple distribution types, perturbation analysis, and an interactive GUI.
Documentation, comments and Unit tests written by AI.

## Table of Contents
- [Features](#features)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Directory Structure](#directory-structure)
- [Usage](#usage)
- [Distribution Types](#distribution-types)
- [GUI Guide](#gui-guide)
- [Command Line Usage](#command-line-usage)
- [File Formats](#file-formats)

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



## Release v3 (2026-01-28)

- Repository cleanup: removed legacy in-memory pickle flows and temporary debug tools.
- Added deterministic per-MM sampling for presets and CSV/Excel parity in generation.
- Renamed `sensivitiy` → `sensitivity` and updated Quickstart and tests.
- Added basic unit tests and a pytest configuration to streamline CI.


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
   - Go to "Export" tab
   - Choose export mode (new file or update current)
   - Click "Export to Excel"
   - Files are saved in `Distributions/` folder

### Using Command Line

```bash
python3 main.py -f Distributions/your_file.xlsx
```

If you don’t pass `--output`, a window opens with interactive export shortcuts.
If you pass `--output`, the combined figure is saved to that path and the script exits without opening a window.

## Directory Structure

```
NewAthenaE2EPSF_v2/
├── README.md                      # This file
├── requirements.txt               # Python dependencies
├── distributions_rotated.py       # Core distribution functions
├── main.py                        # CLI analyzer and plotter
├── optimize_mm_rows.py             # Optimizer + placement + Excel writer (preserves formatting)
├── gui_distributions.py           # Interactive GUI application
├── Distributions/                 # Excel spreadsheet location
│   ├── Test_Distribution.xlsx    # Example file
│   └── your_data.xlsx            # Your files go here (add your own)
└── Figures/                       # Exported plots location
    ├── E2E_PSF_*.png
    └── Encircled_Energy_*.png
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
- `-f, --file FILE`: Path to Excel file (default: `Distributions/Test_Distribution.xlsx`)
- `--normalize`: Normalize PSF to unit integral
- `--no-normalize`: Disable normalization
- `--output FILE`: Save combined plot to file (in `Figures/` folder)
- `--mode {coarse,fine,extra-fine}`: Runtime mode. Controls plotting + optimization speed/accuracy.
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

5. **Extra-fine mode** (deeper optimization, up to ~5 minutes):
   ```bash
   python3 main.py -f Distributions/my_data.xlsx --mode extra-fine
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

10. **Optimize with extra-fine mode**:
   ```bash
   python3 main.py -f Distributions/my_data.xlsx --mode extra-fine --optimize
   ```

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
- `--mode {coarse,fine,extra-fine}`: Controls optimization budget and sampling density
- `--optimize`: Enable row-wise MM# assignment optimization

**Output Files**:
- Input: `Distributions/my_data.xlsx`
- Output: `Distributions/my_data_optimised.xlsx`
- Plots: `Figures/E2E_PSF_*_optimised_*.png`

**Note:** Files like `*_optimised.xlsx` and `*_placed.xlsx` are generated outputs and are not shipped with this project. The `Distributions/` folder in the repo only includes `Test_Distribution.xlsx` by default.

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
   - Select an existing file (e.g., `Test_Distribution.xlsx`)
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
│    [Distributions/Test_Distribution.xlsx    ] [Browse...]       │
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
   - Select `Distributions/Test_Distribution.xlsx` (or your file)
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

**Version**: 2.0  
**Date**: January 2026  
**Python**: 3.8+

### Notable changes in v2.0

- Pseudo-Voigt support with independent `alpha_rad` / `alpha_azi`
- Standard preset loading from the workbook (`MM_PSF` preset table starting at cell K1)
- Row-wise MM# optimizer + deterministic placement seeds (`cross`, `x_axis`, `elliptical`)
- Rotation-invariant HEW evaluation on a polar grid
- Figure exports saved into `Figures/` with timestamped filenames

### Runtime mode note

The `--mode {coarse,fine,extra-fine}` flag trades speed vs sampling density for both plotting and optimization. Start with `coarse` to validate inputs and workflow, then switch to `fine` / `extra-fine` for final runs if needed.

---

**License**: Internal Use
