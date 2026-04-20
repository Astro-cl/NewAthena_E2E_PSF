# Feature Documentation — April 2026 Changes

> **Period covered:** 10 – 20 April 2026  
> **Commits:** `84aa143` through `3916a9f` (40+ commits)  
> **Repository:** `NewAthena_E2E_PSF` on GitLab ESA (`origin/main`)

This document provides extensive technical documentation of the major
features introduced during 10–20 April 2026:

1. [Off-Axis Pointing Implementation](#1-off-axis-pointing-implementation)
2. [Defocus Implementation](#2-defocus-implementation)
3. [HEW Degradation Implementation](#3-hew-degradation-implementation)
   - Includes [Energy-Dependent Sigma Scaling](#313-energy-dependent-sigma-scaling) (§3.13)
4. [Batch Combinations Mode](#4-batch-combinations-mode)

---

## 1. Off-Axis Pointing Implementation

### 1.1 Background and Motivation

Previously, off-axis pointing was handled by writing modified values directly
into the Thermal sheet columns (`d_therm_rotx`, `d_therm_roty`). This
co-mingled the off-axis contribution with genuine thermal perturbations,
making it impossible to distinguish the two sources when inspecting an output
workbook.

The new architecture introduces a dedicated **"Extra PSF shifts"** sheet that
cleanly separates off-axis contributions from thermal perturbations.

### 1.2 Key Commits

| Commit | Date | Description |
|--------|------|-------------|
| `84aa143` | 14 Apr | Initial `--batch-combinations` CLI; first off-axis arcmin→arcsec conversion |
| `ef94e88` | 17 Apr | Moved off-axis/defocus to dedicated "Extra PSF shifts" sheet |

### 1.3 Unit Conversion

Off-axis pointing angle is specified in **arcminutes** (user input / batch
combinations file). It represents a single off-axis angle that is decomposed
equally into X and Y rotation components:

```
d_extra_rotx [arcsec] = offaxis_arcmin × 60 / √2
d_extra_roty [arcsec] = offaxis_arcmin × 60 / √2
```

The `× 60` factor converts arcminutes to arcseconds. The `/ √2` factor
distributes the total off-axis angle equally between the two orthogonal axes,
since the combined rotation magnitude is:

$$\sqrt{(\text{rotx})^2 + (\text{roty})^2} = \sqrt{2} \times \frac{\text{offaxis} \times 60}{\sqrt{2}} = \text{offaxis} \times 60$$

which preserves the total rotation in arcseconds.

**Example:** 20 arcmin off-axis → `20 × 60 / √2 ≈ 848.528` arcsec per axis.

### 1.4 "Extra PSF shifts" Sheet Layout

The sheet has the following columns:

| Column | Header | Unit | Description |
|--------|--------|------|-------------|
| A | `Position #` | — | MM position number (1–600) |
| B | `d_extra_rotx [arcsec]` | arcsec | Extra rotation around X axis |
| C | `d_extra_roty [arcsec]` | arcsec | Extra rotation around Y axis |
| D | `d_extra_z [µm]` | µm | Extra defocus displacement |

### 1.5 Writing (Batch Mode)

In the batch-combinations loop (`main.py` ~line 8016–8044), for each
configuration:

1. The off-axis angle is read from column C of the combinations file (arcmin).
2. Converted to arcsec per axis using the formula above.
3. The "Extra PSF shifts" sheet is created if absent, or cleared and
   repopulated if it exists.
4. All 600 positions receive the same off-axis rotation values (uniform
   pointing offset applies to the entire mirror assembly).

```python
extra_rotx_arcsec = float(offaxis_val) * 60.0 / math.sqrt(2.0)
extra_roty_arcsec = float(offaxis_val) * 60.0 / math.sqrt(2.0)

ws_extra.cell(row=1, column=1, value='Position #')
ws_extra.cell(row=1, column=2, value='d_extra_rotx [arcsec]')
ws_extra.cell(row=1, column=3, value='d_extra_roty [arcsec]')
ws_extra.cell(row=1, column=4, value='d_extra_z [µm]')

for i, pos_num in enumerate(thermal_positions):
    row_idx = i + 2
    ws_extra.cell(row=row_idx, column=1, value=pos_num)
    ws_extra.cell(row=row_idx, column=2, value=extra_rotx_arcsec)
    ws_extra.cell(row=row_idx, column=3, value=extra_roty_arcsec)
```

### 1.6 Reading (Loader)

In `load_gaussians_from_excel()` (`main.py` ~line 1096–1115), after loading
thermal/alignment/gravity perturbation sheets:

1. The "Extra PSF shifts" sheet is read via `pd.read_excel()`.
2. Columns are parsed with `pd.to_numeric(..., errors='coerce').fillna(0.0)`.
3. Each row is stored in `extra_by_pos[pos]` as a dict:
   ```python
   extra_by_pos[pos] = {
       'd_extra_rotx': float(row['d_extra_rotx [arcsec]']),
       'd_extra_roty': float(row['d_extra_roty [arcsec]']),
       'd_extra_z':    float(row['d_extra_z [µm]']) * 1e-6,  # µm → m
   }
   ```
4. If the sheet is missing, `extra_by_pos` remains empty and no extra shifts
   are applied (backward compatible).

### 1.7 Application to Rotation Calculation

In `compute_total_rot_polar()` (`main.py` ~line 3282–3284), the extra
rotations are **added** to the total X/Y rotation for each position:

```python
if extra_by_pos and pos in extra_by_pos:
    rtx_total += float(extra_by_pos[pos].get('d_extra_rotx', 0.0) or 0.0)
    rty_total += float(extra_by_pos[pos].get('d_extra_roty', 0.0) or 0.0)
```

This means the off-axis contribution combines **additively** with the
alignment, gravity, and thermal rotation contributions.

### 1.8 Test Coverage

**File:** `tests/integration/test_extra_psf_shifts.py` (373 lines)

**`TestOffAxisRotation`** class — 5 tests:
- `test_extra_rotx_roty_added_to_total`: Verifies that extra rotx/roty
  values appear in the output of `compute_total_rot_polar()`.
- `test_extra_rotations_combine_with_thermal`: Verifies additive
  combination (thermal rotx=5 + extra=100 → total=105).
- `test_zero_offaxis_no_effect`: Verifies no perturbation when extra
  shifts are zero.
- `test_offaxis_20arcmin_conversion`: Verifies the arcmin→arcsec
  conversion formula: `20 × 60 / √2 ≈ 848.528`.
- `test_loader_reads_extra_rotations`: End-to-end test via
  `load_gaussians_from_excel()` to verify the sheet is read correctly.

---

## 2. Defocus Implementation

### 2.1 Background and Motivation

Defocus represents a displacement of the detector along the optical axis
(Z direction). Previously, defocus was injected by modifying `d_therm_z`
in the Thermal sheet, which again co-mingled defocus with thermal Z
perturbations. The new implementation writes defocus to the `d_extra_z`
column of the same "Extra PSF shifts" sheet used by off-axis pointing.

### 2.2 Key Commits

| Commit | Date | Description |
|--------|------|-------------|
| `84aa143` | 14 Apr | Initial defocus mm→µm conversion in batch mode |
| `ef94e88` | 17 Apr | Moved defocus to "Extra PSF shifts" sheet |

### 2.3 Unit Conversion Chain

Defocus travels through three unit conversions:

1. **User input (batch file column E):** millimetres (mm)
2. **Written to sheet:** microns (µm) → `defocus_mm × 1e3`
3. **Loaded by code:** metres (m) → `defocus_µm × 1e-6`

```
defocus [mm] → ×1e3 → defocus [µm] (sheet) → ×1e-6 → defocus [m] (code)
```

**Example:** 2.5 mm defocus → 2500 µm in sheet → 2.5 × 10⁻³ m in code.

### 2.4 Writing (Batch Mode)

In the batch-combinations loop (`main.py` ~line 8041):

```python
defocus_um = float(defocus_val) * 1e3  # mm -> µm
ws_extra.cell(row=row_idx, column=4, value=defocus_um)
```

### 2.5 Reading (Loader)

The loader reads `d_extra_z [µm]` and converts to metres:

```python
'd_extra_z': float(row.get('d_extra_z [µm]', 0.0)) * 1e-6  # µm -> m
```

If the column is missing from the sheet, the value defaults to 0.0 (no
defocus).

### 2.6 Application — Z-Projection to Centroid Shift

The defocus displacement is applied via the geometric Z-projection formula
in the centroid calculation (`main.py` ~line 3010–3016). The total Z
displacement combines **all four sources**:

$$d_{z,\text{total}} = d_{\text{align},z} + d_{\text{grav},z} + d_{\text{therm},z} + d_{\text{extra},z}$$

The Z displacement is then projected onto the focal plane using the MM's
geometric position:

$$\text{dm}_x = d_{z,\text{total}} \times \frac{x_{\text{MM}}}{12 - z_{\text{MM}}}$$

$$\text{dm}_y = d_{z,\text{total}} \times \frac{y_{\text{MM}}}{12 - z_{\text{MM}}}$$

where:
- $x_{\text{MM}}$, $y_{\text{MM}}$, $z_{\text{MM}}$ are the geometric
  coordinates of the mirror module (from "MM configuration" sheet)
- The focal length is 12 m (the constant `12` in the denominator)
- `dm_x` and `dm_y` are added to the centroid position (`mux`, `muy`)

This means defocus produces a radially asymmetric centroid shift: MMs
further from the optical axis are displaced more.

```python
d_z_total = d_align_z + d_grav_z + d_therm_z + d_extra_z
if denominator != 0 and d_z_total != 0:
    dm_x = d_z_total * x_MM / denominator
    dm_y = d_z_total * y_MM / denominator
    new_mux += dm_x
    new_muy += dm_y
```

### 2.7 Test Coverage

**File:** `tests/integration/test_extra_psf_shifts.py` (shared with off-axis)

**`TestDefocusExtraZ`** class — 6 tests:
- `test_defocus_shifts_centroid`: 1 mm defocus with x_MM=1.0, z_MM=0 →
  mux = 1e-3 × 1.0 / 12 = 8.333 × 10⁻⁵ m.
- `test_zero_defocus_no_shift`: Zero defocus → zero centroid change.
- `test_defocus_adds_to_thermal_dz`: Confirms additive combination with
  d_therm_z (500 + 300 µm).
- `test_defocus_adds_to_all_dz_sources`: Combines with all four sources
  (align=100 + grav=200 + therm=300 + extra=400 µm).
- `test_defocus_y_projection`: Projects along both X and Y axes
  (x_MM=0.6, y_MM=0.8).
- `test_multiple_mms_different_positions`: Three MMs at different
  positions (x=1/0/-1) produce correctly signed shifts.

**`TestExtraShiftLoaderEdgeCases`** class — 3 edge-case tests:
- `test_missing_extra_sheet_no_error`: No "Extra PSF shifts" sheet → no
  error, zero shifts.
- `test_extra_sheet_without_z_column`: Sheet present but missing D column
  → no defocus applied.
- `test_defocus_mm_to_um_conversion`: Verifies the mm→µm→m conversion
  chain (2.5 mm → 2500 µm → 2.5 × 10⁻³ m).

---

## 3. HEW Degradation Implementation

### 3.1 Background and Motivation

HEW (Half Energy Width) degradation models the broadening of individual
mirror module PSFs as a function of off-axis angle and energy. Each MM's
position-dependent rotation angle is used to look up a degradation value
(in arcseconds of HEW) from a reference table. This degradation is then
converted to a Gaussian sigma and added in quadrature to the base sigma,
broadening the per-MM PSF.

The implementation mirrors the existing vignetting sheet architecture
(separate rotazi and rotrad sheets) and was developed across three commits.

### 3.2 Key Commits

| Commit | Date | Description |
|--------|------|-------------|
| `e1f3f11` | 17 Apr | Core HEW degradation: sheet reading, interpolation, sigma broadening |
| `c73048b` | 17 Apr | VLOOKUP resolver for D/E; sigma_extra → "Extra PSF degradations"; degraded sigma → I/J |

### 3.3 Sheet Architecture

Two new sheets are supported (names matched case-insensitively):

| Sheet Name | Purpose |
|------------|---------|
| `MM HEW degradation rotazi` | HEW degradation as a function of azimuthal rotation |
| `MM HEW degradation rotrad` | HEW degradation as a function of radial rotation |

Sheet constants in `main.py` (line 128–129):
```python
HEW_DEG_ROT_AZI_CANDIDATES = ('MM HEW degradation rotazi',)
HEW_DEG_ROT_RAD_CANDIDATES = ('MM HEW degradation rotrad',)
```

### 3.4 Sheet Layout

Each HEW degradation sheet has two functional areas:

**Columns A–C (per-position data):**

| Column | Header | Description |
|--------|--------|-------------|
| A | `Position #` | MM position (1–600) |
| B | `HEW degradation (arcsec)` | **Output**: interpolated HEW degradation |
| C | `Selected energy [keV]` | Energy for lookup (cell C2 only) |

**Columns H–K (lookup reference table):**

| Column | Header | Description |
|--------|--------|-------------|
| H | `Row` | Configuration row number (1–15) |
| I | `rotazi/rotrad [arcmin]` | Off-axis angle in arcminutes |
| J | `energy [keV]` | Photon energy |
| K | `HEW degradation [arcsec]` | HEW degradation value |

### 3.5 Processing Pipeline

The HEW degradation pipeline runs inside `load_gaussians_from_excel()` and
consists of five stages:

#### Stage 1: Build Lookup Table

The H–K reference table is read and organized into interpolation series
keyed by `(cfg_row, energy)`. Angle values in column I are converted from
**arcminutes to arcseconds** (`× 60`):

```python
for rr in range(2, ws_hew_vals.max_row + 1):
    cfg_row  = ws_hew_vals.cell(row=rr, column=8).value   # H: Row #
    angle_am = ws_hew_vals.cell(row=rr, column=9).value   # I: angle [arcmin]
    energy   = ws_hew_vals.cell(row=rr, column=10).value  # J: energy [keV]
    hew_val  = ws_hew_vals.cell(row=rr, column=11).value  # K: HEW [arcsec]

    xval = float(angle_am) * 60.0  # arcmin → arcsec
    key = (cfg_row, float(energy))
    hew_series[key]['xs'].append(xval)
    hew_series[key]['ys'].append(float(hew_val))
```

#### Stage 2: Select Energy

The energy for the lookup is determined from cell **C2** of the HEW sheet.
If C2 is empty/NaN, falls back to the energy specified elsewhere in the
workbook (e.g., from the vignetting sheets or A_eff energy column).

```python
c2_val = ws_hew.cell(row=2, column=3).value
if c2_val is not None and not (isinstance(c2_val, float) and np.isnan(c2_val)):
    hew_sheet_energy = float(c2_val)
```

#### Stage 3: Map Positions to Configuration Rows

Each MM position is mapped to a configuration Row # (1–15) via the
"MM configuration" sheet's `Row #` column. This Row # determines which
row of the lookup table is used for interpolation.

#### Stage 4: Interpolate HEW Degradation per Position

For each position:
1. Get the position's **rotation angle** from `rot_rad_map` or
   `rot_azi_map` (absolute value, in arcsec).
2. Find the matching interpolation series for `(cfg_row, energy)`.
3. Use `np.interp()` for linear interpolation between table points.
4. Write the result to **column B** of the sheet.

```python
rot_val = abs(float(rot_map.get(pos_int, 0.0)))
xs_h, ys_h = hew_series[(cfg_row, energy)]
hew_deg_val = float(np.interp(rot_val, xs_h, ys_h))
ws_hew.cell(row=row_idx, column=2, value=hew_deg_val)
```

If the exact energy is not found, the code falls back to the **nearest
available energy** in the lookup table.

#### Stage 5: Store per-Position Results

The interpolated HEW values are stored in dictionaries for later sigma
broadening:

```python
hew_deg_per_pos_rad[pos_int] = hew_deg_val  # from rotrad sheet
hew_deg_per_pos_azi[pos_int] = hew_deg_val  # from rotazi sheet
```

### 3.6 Sigma Broadening Formula

After all positions are processed, the HEW degradation is applied as
Gaussian sigma broadening (`main.py` ~line 3035–3070).

The conversion from HEW (FWHM) to Gaussian sigma uses the standard
relationship:

$$\text{FWHM} = 2\sqrt{2\ln 2}\;\sigma \quad\Rightarrow\quad \sigma = \frac{\text{FWHM}}{2\sqrt{2\ln 2}}$$

The broadened sigma is computed by **quadrature addition**:

$$\sigma_{\text{new}} = \sqrt{\sigma_{\text{base}}^2 + \sigma_{\text{extra}}^2}$$

where:

$$\sigma_{\text{extra}} = \frac{\text{HEW}_{\text{deg}}}{2\sqrt{2\ln 2}} \times \text{arcsec\_to\_m}$$

The constant $2\sqrt{2\ln 2} \approx 2.3548$ converts FWHM to sigma.

The `arcsec_to_m` conversion factor is:

$$\text{arcsec\_to\_m} = \frac{12 \times \pi}{180 \times 3600} \approx 5.818 \times 10^{-5}\;\text{m/arcsec}$$

where `12` is the focal length in metres.

**Directional broadening:**
- `rotrad` HEW degradation broadens `sigma_rad` (radial component)
- `rotazi` HEW degradation broadens `sigma_azi` (azimuthal component)

```python
_fwhm_to_sigma = 2.0 * np.sqrt(2.0 * np.log(2.0))
_arcsec_to_m = 12.0 * np.pi / 180.0 / 3600.0

# Radial broadening
hew_rad = hew_deg_per_pos_rad.get(pos)
if hew_rad is not None and hew_rad > 0:
    sigma_extra = (hew_rad / _fwhm_to_sigma) * _arcsec_to_m
    old_sigma = float(df.at[idx, 'sigma_rad'])
    df.at[idx, 'sigma_rad'] = np.sqrt(old_sigma**2 + sigma_extra**2)

# Azimuthal broadening
hew_azi = hew_deg_per_pos_azi.get(pos)
if hew_azi is not None and hew_azi > 0:
    sigma_extra = (hew_azi / _fwhm_to_sigma) * _arcsec_to_m
    old_sigma = float(df.at[idx, 'sigma_azi'])
    df.at[idx, 'sigma_azi'] = np.sqrt(old_sigma**2 + sigma_extra**2)
```

### 3.7 Excel Writeback — "Extra PSF degradations" Sheet

After broadening, the per-position **sigma_extra** (the added broadening
component, in arcsec) is written to the "Extra PSF degradations" sheet
(`main.py` ~line 3077–3100):

| Column | Header | Description |
|--------|--------|-------------|
| A | `Position #` | MM position (1–600) |
| B | sigma_extra_rad | Extra radial sigma in arcsec |
| C | sigma_extra_azi | Extra azimuthal sigma in arcsec |

These are **sigma values** (not HEW), already divided by `FWHM_TO_SIGMA`.

### 3.8 Excel Writeback — MM_PSF Columns I/J

The **final degraded sigma** per MM is written to MM_PSF columns I and J
(`main.py` ~line 3100–3130):

| Column | Header | Description |
|--------|--------|-------------|
| I | `sigma_rad_deg [arcsec]` | Final degraded radial sigma |
| J | `sigma_azi_deg [arcsec]` | Final degraded azimuthal sigma |

Values are converted from the DataFrame's internal metre representation
back to arcseconds:

```python
_m_to_arcsec = 1.0 / _arcsec_to_m
ws_hew_psf.cell(row=r, column=9,  value=float(row['sigma_rad']) * _m_to_arcsec)
ws_hew_psf.cell(row=r, column=10, value=float(row['sigma_azi']) * _m_to_arcsec)
```

### 3.9 VLOOKUP Resolver for D/E Columns

#### The Problem

MM_PSF columns D and E (`sigma_rad [arcsec]`, `sigma_azi [arcsec]`) contain
**VLOOKUP formulas** in the original workbook:

```
=VLOOKUP(VLOOKUP($A2,'MM configuration'!$A$2:$H$601,3),$M$31:$Q$45,4)
=VLOOKUP(VLOOKUP($A2,'MM configuration'!$A$2:$H$601,3),$M$31:$Q$45,5)
```

When openpyxl opens and saves the workbook, it **strips the cached formula
values**. On subsequent runs, `data_only=True` returns `None` for these
cells, causing `sigma_rad` to become NaN → 0.0 (via `.fillna(0.0)`). This
broke the broadening formula: `sqrt(0² + extra²) = extra` instead of the
correct `sqrt(base² + extra²)`.

#### The Solution — Python VLOOKUP Resolver

A Python-based VLOOKUP fallback was added (`main.py` ~line 534–638) that
activates when D/E values are missing or zero:

1. **Inner VLOOKUP:** Read "MM configuration" sheet: column A (Position #)
   → column C (Row #). This maps each MM to a configuration row (1–15).

2. **Outer VLOOKUP:** Read the preset table in MM_PSF at rows M30:Q45.
   The table has 15 rows for Row # 1–15:
   - Column M (13): Row #
   - Column N (14): width
   - Column O (15): length
   - Column P (16): sigma_rad [arcsec]
   - Column Q (17): sigma_azi [arcsec]

3. The resolver chains both lookups:
   `MM# → Row# → (sigma_rad, sigma_azi)` from the preset table.

4. Resolved values are stored in the DataFrame and also **written back to
   D/E as plain numbers** (replacing the formulas) in both save points.

```python
# Inner lookup: MM# -> Row#
mm_to_rownum = {}
for r in range(2, ws_cfg.max_row + 1):
    pos_v = ws_cfg.cell(row=r, column=1).value  # A: Position #
    row_v = ws_cfg.cell(row=r, column=3).value  # C: Row #
    mm_to_rownum[int(pos_v)] = int(row_v)

# Outer lookup: Row# -> sigma from preset table
rownum_to_sigma = {}
for r in range(preset_header_row + 1, ws_psf.max_row + 1):
    rv = ws_psf.cell(row=r, column=13).value   # M: Row#
    sig_r = ws_psf.cell(row=r, column=16).value  # P: sigma_rad
    sig_a = ws_psf.cell(row=r, column=17).value  # Q: sigma_azi
    rownum_to_sigma[int(rv)] = (float(sig_r), float(sig_a))
```

#### D/E Persistence

To prevent the issue from recurring, D/E values are written back as **plain
float numbers** (replacing any VLOOKUP formulas) in two locations:

1. **First save** (`main.py` ~line 2791–2822): Inside the HEW/A_eff
   processing block, writes `df['sigma_rad [arcsec]']` and
   `df['sigma_azi [arcsec]']` to columns D(4) and E(5) of MM_PSF.

2. **Second save** (`main.py` ~line 3140–3157): Inside the `wb_hew` save
   block (which also writes I/J and Extra PSF degradations), writes the
   same base sigma values again to ensure consistency.

On subsequent runs, D/E contain plain numbers (no VLOOKUP), so no
resolution is needed. The log message `"INFO: resolved D/E VLOOKUP for N
MMs from preset table"` appears only on the first run after the original
formula workbook has been opened/saved by openpyxl.

### 3.10 Preset Table Column Shift (K → M)

The new I/J degraded sigma columns occupy columns 9–10 of MM_PSF. To avoid
conflict, the **preset distribution table** (which was previously read from
column K = index 10) was shifted to column M = index 12:

**File:** `gui_distributions.py` (line 622):
```python
# Before:
start_col = 10  # K is column 11, 0-indexed = 10

# After:
start_col = 12  # M is column 13, 0-indexed = 12
```

### 3.11 Data Flow Summary

```
┌─────────────────────────────┐
│   HEW Degradation Sheets    │
│  (rotazi / rotrad)           │
│  Columns H-K: lookup table   │
└──────────────┬──────────────┘
               │ read & build interpolation series
               ▼
┌─────────────────────────────┐
│  Per-Position Interpolation  │
│  angle (arcsec) → HEW (arcsec) │
│  Write result → column B     │
└──────────────┬──────────────┘
               │ collect per-pos HEW values
               ▼
┌─────────────────────────────┐
│    Sigma Broadening          │
│  σ_new = √(σ_base² + σ_extra²) │
│  σ_extra = HEW / 2√(2·ln2)  │
└──────────────┬──────────────┘
               │
        ┌──────┴──────┐
        ▼             ▼
┌──────────────┐ ┌──────────────┐
│ Extra PSF    │ │ MM_PSF       │
│ degradations │ │ Cols I/J     │
│ Cols B/C     │ │ (degraded σ) │
│ (σ_extra)    │ │              │
└──────────────┘ └──────────────┘
```

### 3.12 Test Coverage

**File:** `tests/integration/test_hew_degradation.py` (399 lines)

**`TestHEWDegradationWriteBack`** class — 6 tests:
- `test_writes_column_b_for_both_sheets`: Column B populated for both
  rotazi and rotrad.
- `test_interpolation_at_zero_angle`: Zero rotation → zero HEW.
- `test_interpolation_midpoint`: 30 arcsec at midpoint of (0,0)–(60,1.2)
  → ~0.6 arcsec.
- `test_no_hew_sheet_no_crash`: Missing sheets → no error.
- `test_energy_selection_from_c2`: Energy from C2 selects correct series
  (1 keV vs 7 keV).
- `test_multiple_cfg_rows`: Different Row # produces different broadening.

**`TestHEWSigmaBroadening`** class — 8 tests:
- `test_positive_hew_broadens_sigma_rad`: Verifies radial broadening with
  2 arcsec HEW at σ₀ = 4 arcsec.
- `test_positive_hew_broadens_sigma_azi`: Same for azimuthal.
- `test_negative_hew_no_broadening`: Negative HEW → no change.
- `test_zero_hew_no_broadening`: Zero HEW → no change.
- `test_sigmax_sigmay_match_broadened`: sigmax/sigmay copies track
  sigma_rad/sigma_azi.
- `test_per_position_mapping`: Different Row # → different broadening
  magnitudes.
- `test_only_azi_sheet_broadens_azi_only`: Only rotazi present → only
  sigma_azi broadened.
- `test_broadening_formula_exact`: Exact numeric verification with
  σ₀ = 5 arcsec, HEW = 3.5 arcsec.

### 3.13 Energy-Dependent Sigma Scaling

#### Background and Motivation

The angle-dependent HEW degradation (sections 3.3–3.6) broadens each MM's
sigma based on its off-axis rotation angle at a single energy. However,
the PSF width of X-ray optics also depends on photon energy: higher energies
(shorter wavelengths) experience progressively larger scattering from mirror
surface roughness and figure errors. To capture this effect, an additional
**multiplicative energy-dependent scaling factor** is applied to all MM
sigmas after any angle-based HEW broadening.

This factor is stored in a dedicated sheet and is applied **unconditionally**
— it runs even when no rotazi/rotrad HEW degradation sheets are present,
allowing energy dependence to be modelled independently of off-axis
degradation.

#### Key Commits

| Commit | Date | Description |
|--------|------|-------------|
| `66dbd4c` | 18 Apr | Energy-dependent sigma scaling feature |
| `f312b78` | 18 Apr | Fix: move energy scaling outside HEW block |
| `bb26046` | 19 Apr | Unit tests for HEW energy-dependent sigma scaling |

#### Sheet: "MM HEW degradation energy"

A new sheet named **"MM HEW degradation energy"** (matched case-insensitively
by checking for `'hew'`, `'degradation'`, and `'energy'` in the sheet name)
contains a two-column table:

| Column | Header | Unit | Description |
|--------|--------|------|-------------|
| A | `Energy [keV]` | keV | Photon energy |
| B | `Sigma scaling factor` | — | Multiplicative factor for sigma |

**Example table (from test data):**

| Energy [keV] | Sigma scaling factor |
|:---:|:---:|
| 0.2 | 0.98 |
| 0.35 | 0.98 |
| 1.0 | 1.00 |
| 2.0 | 1.03 |
| 4.0 | 1.09 |
| 7.0 | 1.18 |
| 10.0 | 1.23 |
| 12.0 | 1.28 |

The factor at 1 keV is 1.0 (reference energy — no scaling). Values > 1.0
indicate increasing PSF broadening at higher energies.

#### Energy Selection

The energy used for the lookup is `sel_energy`, which is determined earlier
in the pipeline from:
1. Cell **D2** of the `A_eff` sheet (e.g., `"7.0 keV"`)
2. Cell **C2** of the vignetting sheets
3. Cell **C2** of the HEW degradation rotazi/rotrad sheets
4. Fallback: 0.2 keV

In batch mode, the energy is set from column D of the combinations file
and propagated to all relevant cells (A_eff D2, vignetting C2, HEW
degradation C2).

#### Processing Pipeline

The energy-dependent scaling runs in `load_gaussians_from_excel()`
(`main.py` ~line 3229) **after** any angle-dependent HEW broadening:

1. **Open workbook** and locate the "MM HEW degradation energy" sheet.
2. **Read columns A and B** into energy/factor arrays.
3. **Sort by energy** and **interpolate** using `np.interp()` at `sel_energy`.
4. **Multiply** both `sigma_rad` and `sigma_azi` by the interpolated factor.

```python
_hew_energy_factor = float(np.interp(float(sel_energy), _he_energies, _he_factors))

if _hew_energy_factor != 1.0:
    df['sigma_rad'] = df['sigma_rad'].astype(float) * _hew_energy_factor
    df['sigma_azi'] = df['sigma_azi'].astype(float) * _hew_energy_factor
```

**Extrapolation behaviour:** `np.interp` clamps to the boundary values —
energies below 0.2 keV use factor 0.98; energies above 12.0 keV use 1.28.

#### Interaction with Angle-Based HEW Broadening

When both angle-dependent broadening and energy scaling are active, the
operations are applied **sequentially**:

$$\sigma_{\text{final}} = \underbrace{\sqrt{\sigma_{\text{base}}^2 + \sigma_{\text{extra}}^2}}_{\text{angle broadening}} \;\times\; f(E)$$

where $f(E)$ is the energy-dependent scaling factor. This means:
- The angle-based broadening adds scatter in quadrature (models
  off-axis aberration)
- The energy scaling multiplies uniformly (models wavelength-dependent
  surface scattering)

#### Writeback to Workbook

After energy scaling, the final sigma values (including the energy factor)
are written to **MM_PSF columns I/J** (`sigma_rad_deg`, `sigma_azi_deg`)
as arcseconds. This ensures the workbook always reflects exactly the sigma
values used in the PSF aggregation.

#### Data Flow Diagram

```
┌────────────────────────────────────────┐
│  "MM HEW degradation energy" Sheet      │
│  Col A: Energy [keV]                     │
│  Col B: Sigma scaling factor             │
└────────────────────┬───────────────────┘
                     │ read & sort
                     ▼
┌────────────────────────────────────────┐
│  np.interp(sel_energy, energies, factors)│
│  → factor f(E)                           │
└────────────────────┬───────────────────┘
                     │
                     ▼
┌────────────────────────────────────────┐
│  σ_rad *= f(E)                           │
│  σ_azi *= f(E)                           │
│  (applied to ALL 600 MMs uniformly)      │
└────────────────────┬───────────────────┘
                     │
                     ▼
┌────────────────────────────────────────┐
│  Write final σ → MM_PSF cols I/J          │
│  Copy to sigmax/sigmay for aggregation   │
└────────────────────────────────────────┘
```

#### Test Coverage

**File:** `tests/integration/test_hew_degradation.py`

**`TestHEWEnergyScaling`** class — 11 tests:
- `test_factor_applied_at_7kev`: At 7 keV factor = 1.18 → σ scaled.
- `test_factor_unity_at_1kev`: At 1 keV factor = 1.0 → σ unchanged.
- `test_interpolation_at_intermediate_energy`: At 5.5 keV →
  linearly interpolated between (4, 1.09) and (7, 1.18).
- `test_no_energy_sheet_no_scaling`: Missing sheet → no scaling applied.
- `test_energy_scaling_applies_without_hew_broadening`: Works
  independently of rotazi/rotrad sheets.
- `test_energy_scaling_combined_with_hew_broadening`: Verifies
  sequential application (broadening then scaling).
- `test_sigmax_sigmay_include_energy_scaling`: sigmax/sigmay (used
  by aggregation) include the energy factor.
- `test_ij_columns_match_dataframe`: Workbook I/J columns contain
  the energy-scaled sigma values.
- `test_extrapolation_below_range`: Energy below table → clamped to
  lowest factor (0.98).
- `test_extrapolation_above_range`: Energy above table → clamped to
  highest factor (1.28).
- `test_scaling_all_mms`: Factor applied uniformly to all 600 positions.

---

## 4. Batch Combinations Mode

### 4.1 Background and Motivation

The batch mode enables automated execution of the E2E PSF pipeline across
multiple parameter configurations (off-axis angles, energies, defocus
values) defined in a single spreadsheet. Each configuration produces an
independent output workbook and export package. The mode is designed for
headless operation (no GUI, no interactive plots).

### 4.2 Key Commits

| Commit | Date | Description |
|--------|------|-------------|
| `84aa143` | 14 Apr | Initial `--batch-combinations` CLI and batch loop |
| `43a52c1` | 14 Apr | Batch pytest with sample input and ZIP validation |
| `76fae81` | 14 Apr | Always create per-config ZIP; add tests |
| `4a80ad3` | 15 Apr | Prefer full export folder for ZIP creation |
| `4bea47f` | 15 Apr | Write to canonical thermal columns; skip trailing-underscore headers |
| `f7b6ca2` | 15 Apr | Remove extension from ZIP, append timestamp |
| `bd471cc` | 15 Apr | Rename artifacts and create ZIP with controlled file order |
| `2d2c01b` | 15 Apr | Disable interactive plotting; noop `plt.show()` |
| `9044d39` | 15 Apr | Group ZIPs under `Exports/Export_<timestamp>/` |
| `1ba332f` | 15 Apr | Group ZIPs in `Export_<input>_<ts>` folder |
| `0fae935` | 15 Apr | General cleanup of batch feature |
| `268be3a` | 15 Apr | Fallback read Fit_parameters from modified workbook for aggregation |
| `ef94e88` | 17 Apr | Moved off-axis/defocus to "Extra PSF shifts" sheet |

### 4.3 CLI Interface

```
python main.py --batch-combinations <path_to_combinations.xlsx> \
               --file <base_workbook.xlsx>
```

**Arguments:**

| Argument | Description |
|----------|-------------|
| `--batch-combinations` | Path to Excel file with parameter combinations |
| `--file` / `-f` | Path to the base input workbook |

### 4.4 Combinations File Format

The combinations Excel file has one configuration per row:

| Column | Index | Header | Unit | Description |
|--------|-------|--------|------|-------------|
| A | 0 | (ignored) | — | Row number or ID |
| B | 1 | prefix | — | Configuration name (used in output naming) |
| C | 2 | off-axis | arcmin | Off-axis pointing angle |
| D | 3 | energy | keV | Photon energy |
| E | 4 | defocus | mm | Defocus displacement |
| F | 5 | run_mode | — | Optional: `fine` or `coarse` |

**Example:**

| A | B | C | D | E |
|---|---|---|---|---|
| 1 | On-axis_1keV | 0 | 1.0 | 0.0 |
| 2 | Off5_1keV | 5.0 | 1.0 | 0.0 |
| 3 | Off10_1keV_def1 | 10.0 | 1.0 | 1.0 |

### 4.5 Per-Configuration Processing

For each row in the combinations file (`main.py` ~line 7960–8044):

1. **Copy base workbook** with prefix-based naming:
   `<prefix>_<base_filename>.xlsx`

2. **Modify Thermal energy columns:**
   Write the configuration energy to all thermal energy columns.

3. **Create/populate "Extra PSF shifts" sheet:**
   - Off-axis: `offaxis_arcmin × 60 / √2` → `d_extra_rotx` and `d_extra_roty`
   - Defocus: `defocus_mm × 1000` → `d_extra_z [µm]`

4. **Set vignetting/HEW energy:**
   Write energy to cell C2 of vignetting and HEW degradation sheets.

5. **Disable interactive plotting:**
   `matplotlib.use('Agg')` and `plt.show = lambda *a, **k: None`

6. **Run main analysis pipeline** on the modified workbook.

7. **Create export package** with `--export-package` semantics.

### 4.6 Export Packaging and ZIP Creation

Each configuration produces a ZIP export (`main.py` ~line 8689–8755):

1. Prefer zipping the **full export folder** created by `--export-package`
   (includes FITS, PNG, workbook, fit parameters).
2. Fallback: ZIP just the modified workbook.
3. ZIP naming: `<prefix>_<base_stem>_<YYYYMMDD_HHMMSS>.zip`
4. All ZIPs are grouped under `Exports/Export_<input_stem>_<timestamp>/`.

### 4.7 Aggregated Results

After all configurations are processed, an **aggregated results** workbook
is created containing one row per configuration with key metrics:

| Field | Description |
|-------|-------------|
| `configuration_number` | Sequential index |
| `configuration_name` | From column B of combinations file |
| `offaxis_arcmin` | Input off-axis value |
| `energy_keV` | Input energy |
| `defocus_mm` | Input defocus |
| `HEW` | Computed HEW for this configuration |
| `EEF50`, `EEF80`, `EEF90` | EEF diameter metrics |
| `Aeff_loss` | Fractional A_eff loss |
| Fit parameters | King / pseudo-Voigt / Pearson4 fit results |

### 4.8 Headless Operation

The batch mode ensures no GUI interaction:

- `matplotlib.use('Agg')` backend is set before any plotting
- `plt.show()` is replaced with a no-op lambda
- No Tkinter windows are created
- The `--suppress-output` flag can be used to reduce console noise

### 4.9 Test Coverage

**File:** `tests/test_batch_combinations.py` (78 lines)

- `test_batch_combinations_creates_zips`: Creates a minimal 3-configuration
  combinations file (cfgA, cfgB, cfgC) with a sample workbook containing
  Thermal, Vignetting, MM_PSF, MM configuration, and A_eff sheets. Runs
  the batch CLI as a subprocess and verifies:
  - Exit code is 0
  - `Exports/` directory is created
  - A ZIP file exists for each configuration prefix matching the pattern
    `**/<prefix>_<base_stem>_*.zip`

---

## 5. Summary of All Changes

### 5.1 Files Modified

| File | Lines Changed | Description |
|------|---------------|-------------|
| `main.py` | +1100 / -140 | Core changes: batch loop, Extra PSF shifts read/write, HEW degradation pipeline, energy-dependent sigma scaling, VLOOKUP resolver, sigma broadening, I/J writeback, D/E persistence, batch parallelization |
| `gui_distributions.py` | +5 / -4 | Preset table column shift K → M |
| `tests/integration/test_extra_psf_shifts.py` | +373 (new) | Off-axis and defocus integration tests |
| `tests/integration/test_hew_degradation.py` | +630 (new) | HEW degradation interpolation, broadening, and energy scaling tests |
| `tests/test_batch_combinations.py` | +78 (new) | Batch CLI end-to-end test |
| `.gitignore` | +6 | Exclude `*.zip` and batch artifacts |

### 5.2 New Excel Sheets

| Sheet Name | Purpose |
|------------|---------|
| `Extra PSF shifts` | Off-axis rotations (B,C) and defocus (D) per position |
| `Extra PSF degradations` | Sigma broadening per position (B:rad, C:azi) |
| `MM HEW degradation rotazi` | HEW lookup table and per-position results (azimuthal) |
| `MM HEW degradation rotrad` | HEW lookup table and per-position results (radial) |
| `MM HEW degradation energy` | Energy vs sigma scaling factor table for energy-dependent broadening |

### 5.3 New/Modified MM_PSF Columns

| Column | Content | Status |
|--------|---------|--------|
| D | `sigma_rad [arcsec]` (base) | Persisted as plain numbers (was VLOOKUP) |
| E | `sigma_azi [arcsec]` (base) | Persisted as plain numbers (was VLOOKUP) |
| I | `sigma_rad_deg [arcsec]` | **New**: final degraded sigma |
| J | `sigma_azi_deg [arcsec]` | **New**: final degraded sigma |
| M–Q | Preset distribution table | **Shifted** from K–O to avoid I/J conflict |

### 5.4 Test Suite

82 tests total, all passing. The new features added **31 integration tests**:
- 5 off-axis rotation tests
- 6 defocus tests
- 3 loader edge-case tests
- 14 HEW degradation tests (6 interpolation + 8 broadening)
- 11 energy-dependent sigma scaling tests
- 1 batch combinations end-to-end test

### 5.5 Timeline

```
14 Apr  ┤ Batch combinations (--batch-combinations CLI)
        ├ Off-axis/defocus in Thermal columns (initial)
        ├ Batch ZIP packaging iterative fixes (8 commits)
15 Apr  ┤ Batch packaging finalization (folder grouping, timestamps)
        ├ Aggregated results with fit parameters
        ├ Performance: coarse mode default, skip Pearson4
        ├ A_eff formula evaluation fixes
16 Apr  ┤ A_eff aggregation, PSF export, extra-fine removal
        ├ Documentation updates
17 Apr  ┤ Vignetting sheet name support
        ├ A_eff parsing refinement
        ├ Off-axis/defocus → "Extra PSF shifts" sheet (architecture change)
        ├ HEW degradation sheets (rotazi/rotrad) + sigma broadening
        ├ VLOOKUP resolver, D/E persistence, I/J + Extra PSF degradations
```
