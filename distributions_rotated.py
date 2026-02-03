"""Rotated 2D Gaussian and Pseudo-Voigt utilities.

This module provides helpers to load separable PSF matrices and evaluate
rotated 2D distributions used across the project.

Key functions:
- `gaussian_2d_rotated`: standard rotated 2D Gaussian.
- `pseudo_voigt_2d_rotated`: separable 2D pseudo-Voigt built from two 1D PVs.
- `load_psf_matrix_excel` / `eval_psf_matrix_rotated`: helpers to load and
    resample discrete PSF matrices defined on an (x,y) grid.

All distances are handled in meters in the public APIs; callers convert
arcsec->meters using the project's `arcsec_to_m` factor before calling
matrix loaders where appropriate.
"""
import numpy as np


def _sort_axis_and_reorder(axis: np.ndarray, data: np.ndarray, axis_is_x: bool) -> tuple[np.ndarray, np.ndarray]:
    """Sort a 1D axis and reorder 2D data accordingly.

    axis_is_x=True means axis corresponds to columns of data.
    axis_is_x=False means axis corresponds to rows of data.
    """
    axis = np.asarray(axis, dtype=float)
    order = np.argsort(axis)
    axis_sorted = axis[order]
    if axis_is_x:
        data_sorted = data[:, order]
    else:
        data_sorted = data[order, :]
    return axis_sorted, data_sorted




def load_psf_matrix_excel(
    path: str,
    *,
    arcsec_to_m: float,
) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """Load a PSF matrix from an Excel file.

    Expected format:
    - First row (from column B onward): x pixel positions
    - First column (from row 2 onward): y pixel positions
    - Remaining cells: flux samples on the (y, x) grid

    Pixel positions are assumed to be in arcsec and converted to meters using arcsec_to_m.
    The flux is normalized so that the 2D integral over x,y is 1.
    """
    import pandas as pd

    raw = pd.read_excel(path, header=None, engine="openpyxl")
    if raw.shape[0] < 3 or raw.shape[1] < 3:
        raise ValueError(f"PSF matrix file too small: {path}")

    x = pd.to_numeric(raw.iloc[0, 1:], errors="coerce").to_numpy(dtype=float)
    y = pd.to_numeric(raw.iloc[1:, 0], errors="coerce").to_numpy(dtype=float)
    z_df = raw.iloc[1:, 1:]
    z = z_df.apply(pd.to_numeric, errors="coerce").fillna(0.0).to_numpy(dtype=float)

    # Drop NaN coords by masking columns/rows.
    x_mask = np.isfinite(x)
    y_mask = np.isfinite(y)
    x = x[x_mask]
    y = y[y_mask]
    z = z[np.ix_(y_mask, x_mask)]

    if x.size < 2 or y.size < 2:
        raise ValueError(f"PSF matrix must have >=2 x and >=2 y samples: {path}")

    # Convert to meters.
    x_m = x * float(arcsec_to_m)
    y_m = y * float(arcsec_to_m)

    # Sort axes and reorder Z.
    x_m, z = _sort_axis_and_reorder(x_m, z, axis_is_x=True)
    y_m, z = _sort_axis_and_reorder(y_m, z, axis_is_x=False)

    # Normalize integral to 1 using trapezoidal rule.
    # z shape: (ny, nx)
    z = np.asarray(z, dtype=float)
    z = np.where(np.isfinite(z), z, 0.0)
    integral_x = np.trapz(z, x_m, axis=1)
    integral = float(np.trapz(integral_x, y_m))
    if not np.isfinite(integral) or integral <= 0.0:
        raise ValueError(f"PSF matrix integral is not positive/finite ({integral}) for file: {path}")
    z_norm = z / integral
    return x_m, y_m, z_norm


def eval_psf_matrix_rotated(
    X: np.ndarray,
    Y: np.ndarray,
    *,
    mux: float,
    muy: float,
    theta_deg: float,
    x_axis: np.ndarray,
    y_axis: np.ndarray,
    flux: np.ndarray,
) -> np.ndarray:
    """Evaluate a separable-grid PSF matrix on (X,Y) with translation and rotation.

    The PSF matrix is defined on axes (x_axis, y_axis) in meters with flux[y,x].
    We translate the PSF center to (mux, muy) and rotate it by theta_deg.
    Out-of-bounds samples are treated as 0.
    """
    X = np.asarray(X, dtype=float)
    Y = np.asarray(Y, dtype=float)

    dx = X - float(mux)
    dy = Y - float(muy)
    th = np.deg2rad(float(theta_deg))
    c, s = np.cos(th), np.sin(th)

    # Rotate coordinates into the PSF's intrinsic frame.
    # This matches the rotation convention used in gaussian_2d_rotated.
    u = c * dx + s * dy
    v = -s * dx + c * dy

    x = np.asarray(x_axis, dtype=float)
    y = np.asarray(y_axis, dtype=float)
    z = np.asarray(flux, dtype=float)
    nx = x.size
    ny = y.size

    out = np.zeros_like(u, dtype=float)
    if nx < 2 or ny < 2:
        return out

    # Bounds mask
    inb = (u >= x[0]) & (u <= x[-1]) & (v >= y[0]) & (v <= y[-1])
    if not np.any(inb):
        return out

    uq = u[inb]
    vq = v[inb]

    # Indices
    xi = np.searchsorted(x, uq, side="right") - 1
    yi = np.searchsorted(y, vq, side="right") - 1
    xi = np.clip(xi, 0, nx - 2)
    yi = np.clip(yi, 0, ny - 2)

    x0 = x[xi]
    x1 = x[xi + 1]
    y0 = y[yi]
    y1 = y[yi + 1]

    # Avoid division by zero for duplicate axis values.
    tx = np.where(x1 != x0, (uq - x0) / (x1 - x0), 0.0)
    ty = np.where(y1 != y0, (vq - y0) / (y1 - y0), 0.0)

    z00 = z[yi, xi]
    z10 = z[yi, xi + 1]
    z01 = z[yi + 1, xi]
    z11 = z[yi + 1, xi + 1]

    zq = (
        (1.0 - tx) * (1.0 - ty) * z00
        + tx * (1.0 - ty) * z10
        + (1.0 - tx) * ty * z01
        + tx * ty * z11
    )

    out[inb] = zq
    return out


def gaussian_2d_rotated(
    x, y,
    mux, muy,
    sigmax, sigmay,
    theta,
    amplitude=1.0,
    normalize=False,
    degrees=False,
):
    """
    Rotated 2D Gaussian with principal sigmas and rotation angle.

    Parameters
    ----------
    x, y : array-like
        Arrays (1D or 2D) broadcastable to a grid. Typically meshgrid outputs.
    mux, muy : float
        Center (mean) along x and y.
    sigmax, sigmay : float
        Standard deviations along principal axes (> 0).
    theta : float
        Rotation angle. If degrees=True it is in degrees, otherwise radians.
    amplitude : float
        Overall multiplier. If normalize=True, this scales a normalized PDF.
    normalize : bool
        If True, integral is 1 (before applying amplitude) via 1/(2πσxσy).
    degrees : bool
        Interpret theta in degrees if True.

    Returns
    -------
    np.ndarray
        Gaussian values evaluated on the broadcasted grid of (x, y).
    """
    x = np.asarray(x)
    y = np.asarray(y)

    if sigmax <= 0 or sigmay <= 0:
        raise ValueError("sigmax and sigmay must be > 0")

    th = np.deg2rad(theta) if degrees else theta  # No negation
    c, s = np.cos(th), np.sin(th)

    invsx2 = 1.0 / (sigmax ** 2)
    invsy2 = 1.0 / (sigmay ** 2)
    a = c**2 * invsx2 + s**2 * invsy2          # dx^2 coeff
    b = s * c * (invsx2 - invsy2)              # 2*dx*dy coeff /2
    ccoef = s**2 * invsx2 + c**2 * invsy2      # dy^2 coeff

    dx = x - mux
    dy = y - muy
    exponent = -0.5 * (a * dx**2 + 2.0 * b * dx * dy + ccoef * dy**2)

    coeff = (1.0 / (2.0 * np.pi * sigmax * sigmay)) if normalize else 1.0
    return amplitude * coeff * np.exp(exponent)


def pseudo_voigt_2d_rotated(
    azi, rad,
    muazi, murad,
    sigmaazi, sigmarad,
    theta,
    eta=0.5,
    amplitude=1.0,
    normalize=False,
    degrees=False,
    alphaazi=None,
    alpharad=None,
):
    """
    Rotated 2D Pseudo-Voigt formed by multiplying two 1D Pseudo-Voigt profiles.
    
    Each 1D Pseudo-Voigt is: PV(u) = α*L(u) + (1-α)*G(u)
    where G(u) = exp(-u²/(2σ²)) and L(u) = 1/(1 + u²/σ²)
    
    The 2D shape is: PV_2D(azi,rad) = PV_azi(azi') * PV_rad(rad')
    where (azi', rad') are coordinates in the rotated frame.
    
    Parameters
    ----------
    azi, rad : array-like
        Arrays (1D or 2D) broadcastable to a grid. Typically meshgrid outputs.
    muazi, murad : float
        Center (mean) along azimuthal and radial directions.
    sigmaazi, sigmarad : float
        Width parameters along principal axes (> 0).
    theta : float
        Rotation angle. If degrees=True it is in degrees, otherwise radians.
    eta : float
        Mixing parameter (0 ≤ η ≤ 1) for both axes when alphaazi/alpharad not specified.
        η=0 gives pure Gaussian, η=1 gives pure Lorentzian.
    amplitude : float
        Overall multiplier. If normalize=True, this scales a normalized PDF.
    normalize : bool
        If True, approximate normalization is applied.
    degrees : bool
        Interpret theta in degrees if True.
    alphaazi : float, optional
        Mixing parameter for azimuthal axis (0 ≤ α ≤ 1). If None, uses eta.
    alpharad : float, optional
        Mixing parameter for radial axis (0 ≤ α ≤ 1). If None, uses eta.
    
    Returns
    -------
    np.ndarray
        Pseudo-Voigt values evaluated on the broadcasted grid of (azi, rad).
    """
    azi = np.asarray(azi)
    rad = np.asarray(rad)
    
    if sigmaazi <= 0 or sigmarad <= 0:
        raise ValueError("sigmaazi and sigmarad must be > 0")
    
    # Use alphaazi/alpharad if provided, otherwise default to eta
    aazi = eta if alphaazi is None else alphaazi
    arad = eta if alpharad is None else alpharad
    
    if not (0 <= aazi <= 1):
        raise ValueError("alphaazi must be between 0 and 1")
    if not (0 <= arad <= 1):
        raise ValueError("alpharad must be between 0 and 1")
    
    th = np.deg2rad(theta) if degrees else theta
    c, s = np.cos(th), np.sin(th)
    
    # Rotate coordinates to principal axes
    dazi = azi - muazi
    drad = rad - murad
    azi_rot = c * dazi + s * drad
    rad_rot = -s * dazi + c * drad
    
    # 1D Pseudo-Voigt along azimuthal axis (rotated)
    # Normalized 1D Gaussian: (1/sqrt(2π)) * exp(-u²/2)
    # Normalized 1D Lorentzian: (1/π) * 1/(1 + u²)
    u_azi = azi_rot / sigmaazi
    gaussian_azi = (1.0 / np.sqrt(2.0 * np.pi)) * np.exp(-0.5 * u_azi**2)
    lorentzian_azi = (1.0 / np.pi) / (1.0 + u_azi**2)
    pv_azi = (1 - aazi) * gaussian_azi + aazi * lorentzian_azi
    
    # 1D Pseudo-Voigt along radial axis (rotated)
    u_rad = rad_rot / sigmarad
    gaussian_rad = (1.0 / np.sqrt(2.0 * np.pi)) * np.exp(-0.5 * u_rad**2)
    lorentzian_rad = (1.0 / np.pi) / (1.0 + u_rad**2)
    pv_rad = (1 - arad) * gaussian_rad + arad * lorentzian_rad
    
    # 2D profile is the product of the two 1D profiles
    profile = pv_azi * pv_rad
    
    # Normalization
    if normalize:
        # Each 1D pseudo-Voigt is already normalized to integrate to 1
        # The 2D product integrates to: 1 * 1 * σazi * σrad (from change of variables)
        # So we need to divide by (σazi * σrad)
        coeff = 1.0 / (sigmaazi * sigmarad)
    else:
        coeff = 1.0
    
    return amplitude * coeff * profile
