#!/usr/bin/env python3
"""Compute aggregated E2E PSF from MM_PSF sheet and write to FITS (no astropy required).

Writes a Primary HDU-only FITS file with IEEE64 big-endian data.
"""
from __future__ import annotations
import os
import struct
import time
import subprocess
import numpy as np
import pandas as pd

from distributions_rotated import gaussian_2d_rotated, pseudo_voigt_2d_rotated, load_psf_matrix_excel, eval_psf_matrix_rotated
import main as mainmod


def write_simple_fits(path: str, data: np.ndarray, header_cards: dict | None = None) -> None:
    """Write a minimal Primary HDU FITS file with double precision big-endian data.

    This implements only the subset needed for the project's outputs.
    """
    # Ensure array is 2D (ny, nx)
    arr = np.asarray(data, dtype=np.float64)
    if arr.ndim != 2:
        raise ValueError("data must be 2D")

    ny, nx = arr.shape
    cards = []
    def add_card(key, val, comment=None):
        if isinstance(val, str):
            v = f"'{val}'"
        elif isinstance(val, bool):
            v = 'T' if val else 'F'
        else:
            v = str(val)
        if comment:
            cards.append(f"{key:8s}= {v:20s} / {comment}")
        else:
            cards.append(f"{key:8s}= {v:20s}")

    add_card('SIMPLE', True, 'file does conform to FITS standard')
    add_card('BITPIX', -64, 'number of bits per data pixel')
    add_card('NAXIS', 2, 'number of data axes')
    add_card('NAXIS1', nx, 'length of data axis 1')
    add_card('NAXIS2', ny, 'length of data axis 2')
    add_card('EXTEND', True, 'FITS dataset may contain extensions')
    if header_cards:
        for k, v in header_cards.items():
            add_card(k, v)
    cards.append('END')

    # Pad header to 2880-byte blocks (80-char cards)
    header_str = ''.join(card.ljust(80) for card in cards)
    header_bytes = header_str.encode('ascii')
    # pad to multiple of 2880
    pad = (2880 - (len(header_bytes) % 2880)) % 2880
    header_bytes += b' ' * pad

    # FITS requires big-endian IEEE-754 doubles
    data_be = arr.astype('>f8')
    data_bytes = data_be.tobytes(order='C')
    pad2 = (2880 - (len(data_bytes) % 2880)) % 2880
    data_bytes += b'\x00' * pad2

    with open(path, 'wb') as f:
        f.write(header_bytes)
        f.write(data_bytes)


def compute_and_write(input_xlsx: str = 'Distributions/Test_Distribution.xlsx', sheet: str = 'MM_PSF', out_fits: str | None = None, nx: int = 800, ny: int = 800, normalize: bool = True):
    df = mainmod.load_gaussians_from_excel(input_xlsx, sheet=sheet)
    # Prefer an explicit adjusted A_eff column if present (e.g. accounting for
    # vignetting). Fall back to the loader's `weight` column otherwise.
    # The PSF integral will be scaled to the sum of these adjusted A_eff weights.
    if 'aeff_adjusted' in df.columns:
        raw_w = df['aeff_adjusted'].to_numpy(dtype=float)
    else:
        raw_w = df['weight'].to_numpy(dtype=float)
    # Replace NaNs with 0 (missing A_eff treated as zero)
    raw_w = np.nan_to_num(raw_w, nan=0.0)
    # If there are no entries, fallback to an empty weights array
    weight_arr = raw_w.astype(float)
    total_aeff = float(np.nansum(weight_arr))

    mux = df['mux'].to_numpy(dtype=float)
    muy = df['muy'].to_numpy(dtype=float)
    sigmax = df['sigmax'].to_numpy(dtype=float)
    sigmay = df['sigmay'].to_numpy(dtype=float)
    theta = df['theta_degrees'].to_numpy(dtype=float)
    dist = df['distribution'].astype(str).to_numpy()
    alpha_azi = df.get('alpha_azi', pd.Series(0.5, index=df.index)).to_numpy(dtype=float)
    alpha_rad = df.get('alpha_rad', pd.Series(0.5, index=df.index)).to_numpy(dtype=float)

    # Determine grid bounds similar to plot_sum
    center_x = np.nansum(mux * weight_arr)
    center_y = np.nansum(muy * weight_arr)
    max_sigma = max(np.nanmax(sigmax) if sigmax.size else 0.0, np.nanmax(sigmay) if sigmay.size else 0.0)
    max_center_dist = 0.0
    if mux.size:
        max_center_dist = np.nanmax(np.sqrt((mux - center_x)**2 + (muy - center_y)**2))
    r_max = max_center_dist + 5.0 * max_sigma
    if r_max <= 0:
        r_max = 1e-6

    x = np.linspace(center_x - r_max, center_x + r_max, nx)
    y = np.linspace(center_y - r_max, center_y + r_max, ny)
    X, Y = np.meshgrid(x, y)

    Z = np.zeros_like(X, dtype=float)
    workbook_path = input_xlsx

    for i in range(len(weight_arr)):
        w_i = float(weight_arr[i])
        if w_i == 0:
            continue
        if dist[i].lower() in ('gaussian', 'gauss'):
            Z += w_i * gaussian_2d_rotated(X, Y,
                                           mux=mux[i], muy=muy[i],
                                           sigmax=max(sigmax[i], 1e-12), sigmay=max(sigmay[i], 1e-12),
                                           theta=theta[i], amplitude=1.0, normalize=normalize, degrees=True)
        elif 'pseudo' in dist[i].lower() or 'voigt' in dist[i].lower():
            Z += w_i * pseudo_voigt_2d_rotated(X, Y,
                                               mux[i], muy[i],
                                               sigmax[i], sigmay[i],
                                               theta[i], eta=alpha_rad[i] if not np.isnan(alpha_rad[i]) else 0.5,
                                               amplitude=1.0, normalize=normalize, degrees=True,
                                               alphaazi=alpha_azi[i] if not np.isnan(alpha_azi[i]) else None,
                                               alpharad=alpha_rad[i] if not np.isnan(alpha_rad[i]) else None)
        else:
            # Attempt to resolve custom PSF file stem
            from main import _resolve_custom_psf_path
            p = _resolve_custom_psf_path(workbook_path, dist[i])
            if p:
                try:
                    x_psf, y_psf, f_psf = load_psf_matrix_excel(p, arcsec_to_m=12 * np.pi / 180 / 3600)
                    Z += w_i * eval_psf_matrix_rotated(X, Y, mux=mux[i], muy=muy[i], theta_deg=theta[i], x_axis=x_psf, y_axis=y_psf, flux=f_psf)
                except Exception:
                    # fallback to Gaussian
                    Z += w_i * gaussian_2d_rotated(X, Y,
                                                   mux=mux[i], muy=muy[i],
                                                   sigmax=max(sigmax[i], 1e-12), sigmay=max(sigmay[i], 1e-12),
                                                   theta=theta[i], amplitude=1.0, normalize=normalize, degrees=True)
            else:
                # fallback to Gaussian
                Z += w_i * gaussian_2d_rotated(X, Y,
                                               mux=mux[i], muy=muy[i],
                                               sigmax=max(sigmax[i], 1e-12), sigmay=max(sigmay[i], 1e-12),
                                               theta=theta[i], amplitude=1.0, normalize=normalize, degrees=True)

    # Do NOT renormalize: with normalized per-component PSFs and using
    # raw A_eff weights above, the integral of Z should equal total_aeff.
    integral = np.trapz(np.trapz(Z, x, axis=1), y, axis=0) if Z.size else 0.0


    if out_fits is None:
        ts = time.strftime('%Y%m%d_%H%M%S')
        out_fits = os.path.join('CustomPSFs', f'E2E_aggregated_{ts}.fits')

    os.makedirs(os.path.dirname(out_fits), exist_ok=True)
    # Pixel scales: meters per pixel and arcsec per pixel (x and y)
    pix_m_x = float(x[1] - x[0]) if x.size > 1 else 0.0
    pix_m_y = float(y[1] - y[0]) if y.size > 1 else 0.0
    m_to_arcsec = (180.0 / np.pi) * 3600.0
    pix_as_x = pix_m_x * m_to_arcsec
    pix_as_y = pix_m_y * m_to_arcsec
    # FITS CDELT convention: degrees per pixel
    cdelt1 = pix_as_x / 3600.0
    cdelt2 = pix_as_y / 3600.0

    header = {
        'CREATOR': 'export_e2e_fits.py',
        'DATE': time.strftime('%Y-%m-%dT%H:%M:%SZ'),
        'TOT_AEFF': total_aeff,
        'INTG_Z': float(integral),
        'CDELT1': float(cdelt1),
        'CDELT2': float(cdelt2),
        'PIXAS1': float(pix_as_x),
        'PIXAS2': float(pix_as_y),
        'PIXM1': float(pix_m_x),
        'PIXM2': float(pix_m_y),
    }
    # Try to populate optional metadata from git config if available
    def _git_config(key: str) -> str | None:
        try:
            out = subprocess.check_output(['git', 'config', '--get', key], stderr=subprocess.DEVNULL)
            return out.decode('utf-8').strip()
        except Exception:
            return None

    author = _git_config('user.name') or 'Unknown'
    contact = _git_config('user.email') or 'ivo.ferreira@esa.int'
    orcid = _git_config('user.orcid')
    header['AUTHOR'] = author
    header['CONTACT'] = contact or ''
    header['ORCID'] = orcid or ''
    header['INPUTFN'] = os.path.basename(input_xlsx) if input_xlsx else ''
    write_simple_fits(out_fits, Z, header_cards=header)
    print('Wrote FITS to', out_fits)


if __name__ == '__main__':
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('-i', '--input', default='Distributions/Test_Distribution.xlsx')
    p.add_argument('-s', '--sheet', default='MM_PSF')
    p.add_argument('-o', '--output', default=None)
    p.add_argument('--nx', type=int, default=800)
    p.add_argument('--ny', type=int, default=800)
    p.add_argument('--author', default=None, help='Author name to write into FITS header')
    p.add_argument('--contact', default=None, help='Contact (email) to write into FITS header')
    p.add_argument('--orcid', default=None, help='ORCID identifier to write into FITS header')
    args = p.parse_args()
    # If explicit metadata passed via CLI, inject them into git config locally so compute_and_write can pick them up
    if args.author:
        try:
            subprocess.check_call(['git', 'config', '--local', 'user.name', args.author], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass
    if args.contact:
        try:
            subprocess.check_call(['git', 'config', '--local', 'user.email', args.contact], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass
    if args.orcid:
        try:
            subprocess.check_call(['git', 'config', '--local', 'user.orcid', args.orcid], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass

    compute_and_write(input_xlsx=args.input, sheet=args.sheet, out_fits=args.output, nx=args.nx, ny=args.ny)
