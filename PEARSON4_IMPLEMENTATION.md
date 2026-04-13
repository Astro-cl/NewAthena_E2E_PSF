# Pearson4 Model Fitting Implementation

## Overview
Successfully integrated the Pearson4 distribution model from lmfit alongside the existing modified pseudo-Voigt fit for aggregated PSF analysis.

## Key Features Implemented

### 1. **Pearson4 Model Integration**
- Added `lmfit` package to requirements.txt
- Integrated `Pearson4Model` from lmfit.models
- Pearson4 has 5 parameters:
  - **amplitude**: Overall intensity scale
  - **center**: Peak position [arcsec]
  - **sigma**: Width parameter [arcsec]
  - **expon** (m): Shape exponent
  - **skew** (ν): Skewness parameter

### 2. **Dual Least-Squares Optimization**
- Combined optimization for both intensity profile and cumulative energy (EEF)
- Objective function simultaneously minimizes:
  - Log-intensity residuals (for profile matching)
  - EEF residuals weighted by factor (eef_weight=20.0) for cumulative energy matching
- Uses `scipy.optimize.least_squares` with soft L1 loss function
- Falls back to `curve_fit` if least_squares is unavailable

### 3. **Diagnostic Plots**
- **E2E_fit.png**: Modified pseudo-Voigt fit with log-scale y-axis
  - Data points (black dots)
  - Pseudo-Voigt curve (red line)
  - Log-scale visualization for wide dynamic range

- **E2E_fit_pearson4.png**: New Pearson4 dual-panel plot (log scale)
  - **Left panel**: Intensity fit
    - Data points (black dots)
    - Pearson4 model curve (blue line)
    - Log-scale y-axis
  - **Right panel**: Fit quality metrics
    - Log-residuals (red dots)
    - Reference line at zero
    - Log-scale visualization

### 4. **Parameter Export**
- New **fit_parameters.csv** file generated in `Figures/` directory
- Contains comparison of both fit models:
  - Modified Pseudo-Voigt: A, Gamma_c, Gamma_w, eta, beta, scalar
  - Pearson4: amplitude, center, sigma, expon, skew
- CSV format for easy import to Excel/analysis tools

### 5. **EEF Graph Integration**
- Both fit curves now appear on the E2E Encircled Energy Function plot
- **Pseudo-Voigt fit**: Red dashed line (--) 
- **Pearson4 fit**: Orange dotted line (:)
- Allows visual comparison of fit quality across the full energy range
- 95% percentile clipping for better visibility

## Code Location

### Main Implementation (main.py)
- **Lines ~2850-3150**: Pseudo-Voigt + Pearson4 fitting section
  - Two-stage Gaussian core fit
  - Modified pseudo-Voigt optimization
  - Pearson4 model fitting with combined objectives
  - Diagnostic plot generation
  - Parameter export to CSV

- **Line 3770+**: EEF plot integration
  - Pearson4 curve plotted on E2E graph
  - Proper legend and formatting

### Requirements
- **requirements.txt**: Added `lmfit` dependency

## Usage Example

When running main.py with a PSF distribution file:

```bash
python3 main.py --distributions path/to/distributions.xlsx
```

The following outputs are generated:
1. `Figures/E2E_fit.png` - Pseudo-Voigt intensity fit with log scale
2. `Figures/E2E_fit_pearson4.png` - Pearson4 fit quality (2-panel plot, log scale)
3. `Figures/fit_parameters.csv` - Comparison of fit parameters from both models
4. Updated EEF plot includes both pseudo-Voigt and Pearson4 curves

## Advantages of Multi-Model Approach

1. **Flexible Peak Modeling**: Two complementary models capture different aspects:
   - Pseudo-Voigt: Good for realistic core-wing separation
   - Pearson4: Flexible skewness and shape variations

2. **Robust EEF Matching**: Combined objective ensures smooth cumulative energy distribution

3. **Quality Assessment**: 
   - Diagnostic plots with log scale reveal fit quality across dynamic range
   - Residual plots show systematic deviations
   - EEF comparison shows how well total energy is preserved

4. **Parameter Traceability**: CSV export enables:
   - Model comparison
   - Parameter sensitivity analysis
   - Archival and documentation

## Error Handling

- Graceful fallback if lmfit is not available
- Robust fitting with bounds and initial guesses
- Proper handling of edge cases (empty data, invalid values)
- Optional parameters initialization from pseudo-Voigt results

## Testing

- ✓ Code compiles successfully (syntax validation)
- ✓ Test suite passes (2 tests, 12.42s)
- ✓ Pearson4Model imports and initializes correctly
- ✓ Integration with existing pseudo-Voigt fitting verified

## Future Enhancements

Possible extensions:
- Add more distribution models (exp-gauss, skew-normal, etc.)
- Interactive model selection in GUI
- Automated best-model selection based on AIC/BIC
- Confidence intervals on EEF curves
- Multi-wavelength comparison visualization
