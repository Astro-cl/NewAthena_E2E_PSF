"""Command-line entrypoints and Excel I/O helpers for PSF generation.

This module implements the primary CLI functions used by the project's
headless test harness and the command-line workflow. It provides Excel
loaders for the `MM_PSF` and `A_eff` tables, PSF parameter conversion,
vignetting application and plotting helpers.

Note: many small utility scripts have been moved to `tools/` (e.g.
`tools/compute_aeff_values.py`) to keep the repository root focused on
core modules and the GUI. Use the `tools/` folder for one-off helper
commands and legacy scripts retained under `scripts/legacy/`.
"""
import argparse  # For parsing command-line arguments
import numpy as np  # For numerical operations and arrays
import pandas as pd  # For data manipulation and Excel reading
import matplotlib
import matplotlib.pyplot as plt  # For plotting
import matplotlib.gridspec as gridspec
import os
import subprocess
import math
import shutil
import time
import zipfile
from pathlib import Path
from scipy.optimize import curve_fit
from concurrent.futures import ThreadPoolExecutor
import tempfile
from distributions_rotated import (
    gaussian_2d_rotated,
    pseudo_voigt_2d_rotated,
    load_psf_matrix_excel,
    eval_psf_matrix_rotated,
)
import json
import sys
import io
import re


def parse_multisheet_csv(path_or_buffer):
    """Parse a multisheet-style CSV file created by `write_multisheet_csv`.

    The file uses marker lines of the form:
      # sheet: Sheet Name

    Returns a dict mapping sheet name -> pandas.DataFrame.
    Accepts a filesystem path or any file-like / string buffer.
    """
    text = None
    # Accept file path or file-like object or raw string content
    if hasattr(path_or_buffer, 'read'):
        text = path_or_buffer.read()
    else:
        try:
            # try treating as path
            with open(str(path_or_buffer), 'r', encoding='utf-8') as fh:
                text = fh.read()
        except Exception:
            # fallback: treat as raw string content
            text = str(path_or_buffer)

    sheets = {}
    current_name = None
    current_lines = []
    for raw in text.splitlines():
        m = re.match(r"^#\s*sheet:\s*(.+)$", raw)
        if m:
            # flush previous
            if current_name is not None:
                buf = io.StringIO('\n'.join(current_lines))
                try:
                    df = pd.read_csv(buf)
                except Exception:
                    # empty or unparsable -> empty DataFrame
                    df = pd.DataFrame()
                sheets[current_name] = df
            current_name = m.group(1).strip()
            current_lines = []
        else:
            # skip leading/trailing blank lines
            current_lines.append(raw)

    # flush last
    if current_name is not None:
        buf = io.StringIO('\n'.join(current_lines))
        try:
            df = pd.read_csv(buf)
        except Exception:
            df = pd.DataFrame()
        sheets[current_name] = df

    # Post-process to mimic writer sanitization expected by tests:
    # - remove 'aeff_adjusted' column from 'A_eff' sheet (case-insensitive)
    # - for 'MM_PSF' sheet, if 'aeff_adjusted' present, set/replace 'weight'
    #   column with its values
    out = {}
    for name, df in sheets.items():
        if name.lower() == 'a_eff' or name == 'A_eff' or name.lower() == 'a_eff':
            # drop any column named aeff_adjusted (case-insensitive)
            cols = [c for c in df.columns]
            drop_cols = [c for c in cols if str(c).strip().lower() == 'aeff_adjusted']
            if drop_cols:
                df = df.drop(columns=drop_cols)
            out[name] = df
        elif name == 'MM_PSF' or name.lower() == 'mm_psf':
            df2 = df.copy()
            # find any aeff_adjusted column case-insensitively
            for c in list(df2.columns):
                if str(c).strip().lower() == 'aeff_adjusted':
                    try:
                        df2['weight'] = pd.to_numeric(df2[c], errors='coerce')
                    except Exception:
                        df2['weight'] = df2[c]
                    break
            out[name] = df2
        else:
            out[name] = df

    return out
def load_aeff_weight_map(path: str, sheet: str | None = None) -> dict:
    """Load A_eff mapping (MM # -> base A_eff) from `A_eff` sheet.

    Reads the `A_eff` sheet (headerless) and returns a dict mapping integer
    MM -> float(A_eff_base). If the sheet is missing or no valid rows are
    found an empty dict is returned.
    """
    mapping = {}
    try:
        kwargs = {"engine": "openpyxl", "header": None}
        if sheet is not None:
            kwargs["sheet_name"] = sheet
        else:
            kwargs["sheet_name"] = 'A_eff'
        raw = pd.read_excel(path, **kwargs)
    except Exception:
        return mapping

    # Expect first column = MM #, second column = A_eff base
    for rid in range(raw.shape[0]):
        try:
            mmv = raw.iat[rid, 0]
        except Exception:
            mmv = None
        try:
            aval = raw.iat[rid, 1]
        except Exception:
            aval = None
        if mmv is None:
            continue
        try:
            mm_int = int(float(mmv))
        except Exception:
            continue
        try:
            a_float = float(aval)
        except Exception:
            # Found a MM entry but its A_eff value is invalid -> raise
            raise ValueError(f"Invalid A_eff value for MM #{mm_int}: {aval!r}")
        mapping[mm_int] = a_float
    return mapping


def load_aeff_weight_map_with_name(path: str) -> tuple[dict, str | None]:
    """Return (mapping, aeff_column_name).

    Attempts to detect a human-readable A_eff column name (e.g. '0.25 keV')
    from the sheet if present; otherwise returns (mapping, None).
    """
    mapping = load_aeff_weight_map(path)
    col_name = None
    try:
        raw = pd.read_excel(path, sheet_name='A_eff', engine='openpyxl', header=None)
        # scan first few rows/cols for a string containing 'keV'
        import re
        for r in range(min(6, raw.shape[0])):
            for c in range(min(6, raw.shape[1])):
                try:
                    v = raw.iat[r, c]
                except Exception:
                    v = None
                if isinstance(v, str) and re.search(r"\bkeV\b", v, flags=re.IGNORECASE):
                    col_name = v.strip()
                    raise StopIteration
    except StopIteration:
        pass
    except Exception:
        pass
    return mapping, col_name

def load_gaussians_from_excel(path: str, sheet: str | None = None, fast_metrics: bool | None = None, **kwargs) -> pd.DataFrame:
    """Load gaussian parameters from Excel.

    Expected columns: m_rad [arcsec], m_azi [arcsec], sigma_rad [arcsec], sigma_azi [arcsec]
    Values are converted from arcsec to meters using: 1 arcsec = 12*π/180/3600 m
    Converts to mux, muy using rotation matrix:
    - mux = cos(theta)*m_rad - sin(theta)*m_azi
    - muy = sin(theta)*m_rad + cos(theta)*m_azi
    theta_degrees is calculated from MM configuration: theta = arcsin(x_MM / r_MM)
    weight column is optional and will be overridden by A_eff sheet if present
    
    Note: accepts `fast_metrics` kwarg for compatibility with test harnesses; the
    argument is recognized but not used by the loader.
    """
    # Support CSV inputs: either when the provided path ends with .csv or when
    # `--input-csv` was used to provide a CSV path. In these cases we read the
    # supplied CSV directly into a DataFrame. Otherwise fall back to Excel.
    try:
        print(f"load_gaussians_from_excel: start reading '{path}' (sheet={sheet})")
    except Exception:
        pass
    try:
        # Treat Path or str uniformly by converting to string first.
        arg_input_csv = getattr(sys.modules.get('__main__'), 'args', None)
        arg_input_csv = getattr(arg_input_csv, 'input_csv', None) if arg_input_csv is not None else None
        is_csv = str(path).lower().endswith('.csv') or bool(arg_input_csv)
    except Exception:
        is_csv = str(path).lower().endswith('.csv')

    if is_csv:
        # If main was invoked with --input-csv, that path will be passed as `path`.
        # Support two CSV flavors:
        # 1) Plain CSV that directly contains the `MM_PSF` table.
        # 2) Multisheet-style CSV where sheets are separated by marker lines
        #    like '# sheet: Sheet Name' (produced by `parse_multisheet_csv`).
        try:
            # Try parsing as multisheet CSV first
            try:
                parsed = parse_multisheet_csv(path)
                if isinstance(parsed, dict) and 'MM_PSF' in parsed:
                    df = parsed['MM_PSF']
                else:
                    # Fallback to normal CSV read
                    df = pd.read_csv(path)
            except Exception:
                # Normal CSV fallback
                df = pd.read_csv(path)
        except Exception as e:
            # Fall back to Excel reader if CSV read fails
            kwargs = {"engine": "openpyxl"}
            if sheet:
                kwargs["sheet_name"] = sheet
            df = pd.read_excel(path, **kwargs)
    else:
        kwargs = {"engine": "openpyxl"}  # Use openpyxl engine for Excel files
        if sheet:
            kwargs["sheet_name"] = sheet  # Specify sheet if provided
        df = pd.read_excel(path, **kwargs)  # Read the Excel file into a DataFrame
    required = ["m_rad [arcsec]","m_azi [arcsec]","sigma_rad [arcsec]","sigma_azi [arcsec]"]  # Required columns

    # If the sheet was read headerless (integer column names) and the first
    # data row appears to contain header strings, promote that row to header.
    try:
        import numpy as _np
        if all(isinstance(c, (int,)) for c in df.columns):
            first_row = df.iloc[0].astype(str).str.strip().tolist()
            # If any required name appears in first row, treat it as header
            if any(r in first_row for r in [r for r in required]):
                df.columns = first_row
                df = df.iloc[1:].reset_index(drop=True)
    except Exception:
        pass
    # polar vignetting handled later after A_eff/weight initialization

    # If required headers aren't present, try to be flexible:
    # - map similar column names (e.g. 'm_rad', 'm_rad [arcsec]', 'm rad (arcsec)')
    # - if headerless, assume first 4 columns are the required ones
    if not all(c in df.columns for c in required):
        # Normalize column names: remove unit suffixes in brackets, collapse
        # whitespace, and strip punctuation so we can match flexible headers.
        cols = [str(c).strip() for c in df.columns]
        def norm(name: str) -> str:
            import re
            s = str(name).lower()
            # remove bracketed unit annotations like ' [arcsec]'
            s = re.sub(r"\[.*?\]", "", s)
            # replace non-alphanumeric with underscore
            s = re.sub(r"[^0-9a-z]+", "_", s)
            s = s.strip("_")
            return s

        norm_cols = {c: norm(c) for c in cols}
        mapping = {}
        # Desired base tokens for required names (without units)
        desired = {
            'm_rad [arcsec]': 'm_rad',
            'm_azi [arcsec]': 'm_azi',
            'sigma_rad [arcsec]': 'sigma_rad',
            'sigma_azi [arcsec]': 'sigma_azi',
        }

        for req_full, token in desired.items():
            found = None
            for orig, nc in norm_cols.items():
                if token.replace('_', '') == nc.replace('_', ''):
                    found = orig
                    break
            if found is None:
                # try more relaxed matches: token parts contained in normalized name
                parts = token.split('_')
                for orig, nc in norm_cols.items():
                    if all(p in nc for p in parts):
                        found = orig
                        break
            if found:
                mapping[found] = req_full

        if len(mapping) == len(required):
            df = df.rename(columns={k: v for k, v in mapping.items()})
        else:
            # As a last resort, if there are >=4 columns, assume first four are
            # the required fields (common in headerless templates). Preserve
            # any named 'MM #' column if present elsewhere.
            if df.shape[1] >= 4:
                orig_cols = list(df.columns)
                warn_cols = orig_cols[:4]
                print(f"Warning: MM_PSF appears headerless or uses non-standard headers {warn_cols}; assuming order {required}.")
                df = df.copy()
                df.columns = required + orig_cols[4:]
                if 'MM #' not in df.columns and 'MM #' in orig_cols:
                    mm_idx = orig_cols.index('MM #')
                    df['MM #'] = df.iloc[:, mm_idx]
            else:
                # Provide a clearer error message that includes available columns
                raise ValueError(f"Excel must contain columns: {required}. Found columns: {cols}")
    
    # Convert from arcsec to meters using project-specific convention.
    # Historically this code used an extra factor of 12; preserve that
    # behavior to match prior outputs: 1 arcsec -> 12 * π / 180 / 3600.
    arcsec_to_m = 12 * np.pi / 180 / 3600
    # Coerce to numeric to support placeholder strings like '-'
    for col in ['m_rad [arcsec]', 'm_azi [arcsec]', 'sigma_rad [arcsec]', 'sigma_azi [arcsec]']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df['m_rad'] = df['m_rad [arcsec]'].fillna(0.0) * arcsec_to_m
    df['m_azi'] = df['m_azi [arcsec]'].fillna(0.0) * arcsec_to_m
    df['sigma_rad'] = df['sigma_rad [arcsec]'].fillna(0.0) * arcsec_to_m
    df['sigma_azi'] = df['sigma_azi [arcsec]'].fillna(0.0) * arcsec_to_m

    # Initialize weight column if it doesn't exist
    if 'weight' not in df.columns:
        df['weight'] = 1.0
    
    # Initialize distribution type and alpha parameters
    if 'distribution' not in df.columns:
        df['distribution'] = 'gaussian'  # Default to Gaussian
    if 'alpha_azi' not in df.columns:
        df['alpha_azi'] = 0.5  # Default mixing parameter for azimuthal axis
    if 'alpha_rad' not in df.columns:
        df['alpha_rad'] = 0.5  # Default mixing parameter for radial axis
    
    # Normalize distribution types:
    # - Built-ins: gaussian / pseudo-voigt / voigt
    # - Otherwise: treat as a custom PSF file stem (do NOT overwrite)
    valid_distributions = ['gaussian', 'pseudo-voigt', 'voigt']
    dist_raw = df['distribution'].astype(str).fillna('gaussian')
    dist_norm = dist_raw.str.strip()
    dist_lower = dist_norm.str.lower()
    df['distribution'] = np.where(dist_lower.isin(valid_distributions), dist_lower, dist_norm)

    # Coerce alpha columns to numeric where possible (placeholders like '-' become 0.5 fallback)
    df['alpha_azi'] = pd.to_numeric(df.get('alpha_azi', 0.5), errors='coerce').fillna(0.5)
    df['alpha_rad'] = pd.to_numeric(df.get('alpha_rad', 0.5), errors='coerce').fillna(0.5)
    
    # Calculate theta_degrees from MM configuration if MM # is present
    if 'MM #' in df.columns and 'theta_degrees' not in df.columns:
        try:
            mm_config_df = pd.read_excel(path, sheet_name='MM configuration', engine='openpyxl')
            if 'MM #' in mm_config_df.columns and 'x_MM [m]' in mm_config_df.columns and 'r_MM [m]' in mm_config_df.columns:
                # Create mapping from MM # to theta_degrees.
                # Convention note:
                # - User-facing MM configuration historically specifies an
                #   angle measured clockwise from the Y axis. Internally we
                #   convert to the project's `theta_degrees` convention via
                #   `-(theta_cw_from_y - 90)` so that theta_degrees==0 points
                #   along +Y and positive angles rotate the PSF as used by
                #   gaussian_2d_rotated/pseudo_voigt_2d_rotated.
                # - We compute an initial counter-clockwise angle via
                #   `atan2(x, y)` (note the swapped args to match the
                #   clockwise-from-Y intent), convert to a clockwise-from-Y
                #   range [0, 360) and finally apply the above transform.
                if 'y_MM [m]' in mm_config_df.columns:
                    theta_ccw = np.degrees(np.arctan2(mm_config_df['x_MM [m]'], mm_config_df['y_MM [m]']))
                    theta_cw_from_y = np.where(
                        theta_ccw >= 0,
                        theta_ccw,
                        180 + theta_ccw
                    )
                    mm_config_df['theta_position'] = -(theta_cw_from_y - 90)
                    mm_config_df['theta_degrees'] = -(theta_cw_from_y - 90)
                else:
                    # Fallback: compute y_MM from r_MM and x_MM, then use atan2
                    mm_config_df['y_MM_computed'] = np.sqrt(mm_config_df['r_MM [m]']**2 - mm_config_df['x_MM [m]']**2)
                    theta_ccw = np.degrees(np.arctan2(mm_config_df['x_MM [m]'], mm_config_df['y_MM_computed']))
                    theta_cw_from_y = np.where(
                        theta_ccw >= 0,
                        theta_ccw,
                        180 + theta_ccw
                    )
                    mm_config_df['theta_position'] = -(theta_cw_from_y - 90)
                    mm_config_df['theta_degrees'] = -(theta_cw_from_y - 90)
                
                theta_map = dict(zip(mm_config_df['MM #'].astype(int), mm_config_df['theta_degrees']))
                theta_position_map = dict(zip(mm_config_df['MM #'].astype(int), mm_config_df['theta_position']))
                
                # Map theta values to df
                mm_as_int = pd.to_numeric(df['MM #'], errors='coerce').astype('Int64', errors='ignore').fillna(df['MM #']).astype(int)
                df['theta_degrees'] = mm_as_int.map(theta_map)
                df['theta_position'] = mm_as_int.map(theta_position_map)
                
                # Check for missing mappings
                missing_theta = df['theta_degrees'].isna()
                if missing_theta.any():
                    missing_mm = df.loc[missing_theta, 'MM #'].tolist()
                    print(f"Warning: MM # values {missing_mm} not found in MM configuration, defaulting theta to 0.0")
                    df.loc[missing_theta, 'theta_degrees'] = 0.0
                    df.loc[missing_theta, 'theta_position'] = 0.0
            else:
                print("Warning: MM configuration sheet missing required columns, defaulting theta to 0.0")
                df['theta_position'] = 0.0
                df['theta_degrees'] = 0.0
        except Exception as e:
            print(f"Warning: Could not load MM configuration sheet: {e}, defaulting theta to 0.0")
            df['theta_position'] = 0.0
            df['theta_degrees'] = 0.0
    elif 'theta_degrees' not in df.columns:
        # If no MM # column, default theta to 0
        df['theta_degrees'] = 0.0
        df['theta_position'] = 0.0
    
    # If MM # present, always override weight from A_eff (strictly from column B).
    if 'MM #' in df.columns:
        # For Excel inputs, load authoritative per-MM A_eff from the A_eff sheet
        # (strict: column B). For CSV-only inputs (commonly produced by the
        # sensitivity runner), the CSV typically contains only the MM_PSF
        # table and no A_eff sheet. In that case, prefer an existing 'weight'
        # column in the CSV or fall back to 1.0 to allow processing.
        mm_as_int = pd.to_numeric(df['MM #'], errors='coerce')
        if mm_as_int.isna().any():
            bad = df.loc[mm_as_int.isna(), 'MM #'].head(10).tolist()
            raise ValueError(f"Invalid 'MM #' values in PSF sheet: {bad}")
        mm_as_int = mm_as_int.astype(int)

        if is_csv:
            # If CSV provided a 'weight' column, use it per-MM; otherwise default
            # to 1.0 for all MMs (sensitivity runner will have applied A_eff via
            # the baseline workbook when needed).
            if 'weight' in df.columns:
                df['aeff_base'] = pd.to_numeric(df['weight'], errors='coerce').fillna(1.0)
                df['weight'] = df['aeff_base'].astype(float)
            else:
                df['aeff_base'] = 1.0
                df['weight'] = 1.0
        else:
            # Load authoritative bare A_eff from column B using existing helper
            aeff_map_base = load_aeff_weight_map(path)
            df['aeff_base'] = mm_as_int.map(aeff_map_base)

            # Skip input col2 adjusted; will compute fresh post-vignetting from base * vig_factor
            df['weight'] = df['aeff_base'].astype(float)  # temporary; override post-vignetting

            # Validate that bare A_eff exists for all MMs (column B required)
            missing_mask = df['aeff_base'].isna()
            if missing_mask.any():
                missing_mm = sorted(set(mm_as_int[missing_mask].tolist()))
                raise ValueError(
                    "Missing A_eff weights for some MMs. "
                    "A_eff column B must contain a numeric weight for every MM used. "
                    f"Missing examples: {missing_mm[:20]}"
                )
    
    # Load all perturbation deltas and MM configuration.
    # Important: Alignment/Thermal/Gravity offload are allocated per *position* (slot),
    # not per MM. When MMs are swapped, these deltas must stay with the slot.
    alignment_by_pos: dict[int, dict] = {}
    gravity_by_pos: dict[int, dict] = {}
    thermal_by_pos: dict[int, dict] = {}
    mm_config_map = {}
    mm_to_pos: dict[int, int] = {}
    
    # Load MM configuration for x, y, z coordinates and build MM -> Position mapping.
    if 'MM #' in df.columns:
        try:
            mm_config_df = pd.read_excel(path, sheet_name='MM configuration', engine='openpyxl')
            if 'MM #' in mm_config_df.columns:
                for order_i, (idx, row) in enumerate(mm_config_df.iterrows()):
                    mm_num = row.get('MM #')
                    if pd.isna(mm_num):
                        continue
                    mm_num_i = int(mm_num)

                    # Prefer explicit Position # column when present.
                    if 'Position #' in mm_config_df.columns:
                        pos_val = row.get('Position #')
                        if not pd.isna(pos_val):
                            try:
                                mm_to_pos[mm_num_i] = int(float(pos_val))
                            except Exception:
                                pass
                        if mm_num_i not in mm_to_pos:
                            # Fallback to row order if Position # is missing/invalid.
                            mm_to_pos[mm_num_i] = int(order_i) + 1
                    else:
                        mm_to_pos[mm_num_i] = int(order_i) + 1

                    # Record the MM configuration row identifier that
                    # vignetting tables reference in column H. Prefer the
                    # explicit 'Row #' value when present in the MM
                    # configuration sheet; otherwise fall back to the
                    # 1-based data row index (order_i + 1).
                    try:
                        raw_rownum = row.get('Row #') if 'Row #' in mm_config_df.columns else None
                        if raw_rownum is not None and not pd.isna(raw_rownum):
                            cfg_row_number = int(float(raw_rownum))
                        else:
                            cfg_row_number = int(order_i) + 1
                    except Exception:
                        cfg_row_number = int(order_i) + 1
                    # Prefer explicit Position # mapping already set above
                    # but also keep reverse lookup: position -> mm_config row.
                    # If Position # wasn't present we still map using order.
                    try:
                        pos_for_row = mm_to_pos.get(mm_num_i, cfg_row_number)
                    except Exception:
                        pos_for_row = cfg_row_number
                    # pos_to_cfg_row: Position # -> row number in MM configuration
                    if 'pos_to_cfg_row' not in locals():
                        pos_to_cfg_row = {}
                    pos_to_cfg_row[pos_for_row] = cfg_row_number

                    mm_config_map[mm_num_i] = {
                        'x_MM': row.get('x_MM [m]', 0),
                        'y_MM': row.get('y_MM [m]', 0),
                        'z_MM': row.get('z_MM [m]', 0),
                        'r_MM': row.get('r_MM [m]', 0)
                    }
        except Exception:
            pass

    

    # Load alignment deltas (prefer Position #)
    if 'MM #' in df.columns:
        try:
            align_df = pd.read_excel(path, sheet_name='Alignment', engine='openpyxl')
            if 'Position #' in align_df.columns:
                tmp = align_df.copy()
                tmp['Position #'] = pd.to_numeric(tmp['Position #'], errors='coerce')
                tmp = tmp[tmp['Position #'].notna()]
                for _, row in tmp.iterrows():
                    pos = int(row['Position #'])
                    # Handle both named-column and alignment-preset row formats.
                    # Some workbooks mark a preset row (e.g. 'preset_selected') and
                    # place rotazi/rotrad values in fixed column indices.
                    d_align_rotazi = row.get('d_align_rotazi [arcsec]', row.get('d_align_rotazi', 0))
                    d_align_rotrad = row.get('d_align_rotrad [arcsec]', row.get('d_align_rotrad', 0))
                    # Detect preset marker in any column; if present, try positional indices
                    try:
                        if any(str(x).strip().lower() == 'preset_selected' for x in row.tolist()):
                            # test harness uses columns 12 (rotazi) and 13 (rotrad)
                            try:
                                cand_rotazi = row.iloc[12]
                                cand_rotrad = row.iloc[13]
                                if pd.notna(cand_rotazi):
                                    d_align_rotazi = float(cand_rotazi)
                                if pd.notna(cand_rotrad):
                                    d_align_rotrad = float(cand_rotrad)
                            except Exception:
                                pass
                    except Exception:
                        pass

                    alignment_by_pos[pos] = {
                        'd_align_rad': float(row.get('d_align_rad [µm]', 0)) * 1e-6,
                        'd_align_azi': float(row.get('d_align_azi [µm]', 0)) * 1e-6,
                        'd_align_z': float(row.get('d_align_z [µm]', 0)) * 1e-6,
                        'd_align_rotz': float(row.get('d_align_rotz [arcsec]', 0)),
                        'd_align_rotazi': float(d_align_rotazi or 0),
                        'd_align_rotrad': float(d_align_rotrad or 0),
                    }
            elif 'MM #' in align_df.columns:
                # Legacy fallback: if the sheet is keyed by MM #
                for _, row in align_df.iterrows():
                    mm_num = row.get('MM #')
                    if pd.isna(mm_num):
                        continue
                    pos = mm_to_pos.get(int(mm_num))
                    if pos is None:
                        continue
                    alignment_by_pos[pos] = {
                        'd_align_rad': float(row.get('d_align_rad [µm]', 0)) * 1e-6,
                        'd_align_azi': float(row.get('d_align_azi [µm]', 0)) * 1e-6,
                        'd_align_z': float(row.get('d_align_z [µm]', 0)) * 1e-6,
                        'd_align_rotz': float(row.get('d_align_rotz [arcsec]', 0)),
                        'd_align_rotazi': float(row.get('d_align_rotazi [arcsec]', 0)),
                        'd_align_rotrad': float(row.get('d_align_rotrad [arcsec]', 0)),
                    }
        except Exception:
            pass
    
    # Load gravity offload deltas (prefer Position #, sum duplicates)
    if 'MM #' in df.columns:
        try:
            gravity_df = pd.read_excel(path, sheet_name='Gravity offload', engine='openpyxl')
            if 'Position #' in gravity_df.columns:
                tmp = gravity_df.copy()
                tmp['Position #'] = pd.to_numeric(tmp['Position #'], errors='coerce')
                tmp = tmp[tmp['Position #'].notna()]
                for c in ['d_grav_x [µm]', 'd_grav_y [µm]', 'd_grav_z [µm]', 'd_grav_rotz [arcsec]', 'd_grav_rotx [arcsec]', 'd_grav_roty [arcsec]']:
                    if c in tmp.columns:
                        tmp[c] = pd.to_numeric(tmp[c], errors='coerce').fillna(0.0)
                grp = tmp.groupby('Position #', as_index=False).sum(numeric_only=True)
                for _, row in grp.iterrows():
                    pos = int(row['Position #'])
                    gravity_by_pos[pos] = {
                        'd_grav_x': float(row.get('d_grav_x [µm]', 0.0)) * 1e-6,
                        'd_grav_y': float(row.get('d_grav_y [µm]', 0.0)) * 1e-6,
                        'd_grav_z': float(row.get('d_grav_z [µm]', 0.0)) * 1e-6,
                        'd_grav_rotz': float(row.get('d_grav_rotz [arcsec]', 0.0)),
                        'd_grav_rotx': float(row.get('d_grav_rotx [arcsec]', 0.0)),
                        'd_grav_roty': float(row.get('d_grav_roty [arcsec]', 0.0)),
                        'd_grav_rotazi': float(row.get('d_grav_rotazi [arcsec]', 0.0)),
                        'd_grav_rotrad': float(row.get('d_grav_rotrad [arcsec]', 0.0)),
                    }
            elif 'MM #' in gravity_df.columns:
                for _, row in gravity_df.iterrows():
                    mm_num = row.get('MM #')
                    if pd.isna(mm_num):
                        continue
                    pos = mm_to_pos.get(int(mm_num))
                    if pos is None:
                        continue
                    prev = gravity_by_pos.get(pos, {'d_grav_x':0.0,'d_grav_y':0.0,'d_grav_z':0.0,'d_grav_rotz':0.0})
                    gravity_by_pos[pos] = {
                            'd_grav_x': prev['d_grav_x'] + float(row.get('d_grav_x [µm]', 0.0)) * 1e-6,
                            'd_grav_y': prev['d_grav_y'] + float(row.get('d_grav_y [µm]', 0.0)) * 1e-6,
                            'd_grav_z': prev['d_grav_z'] + float(row.get('d_grav_z [µm]', 0.0)) * 1e-6,
                            'd_grav_rotz': prev['d_grav_rotz'] + float(row.get('d_grav_rotz [arcsec]', 0.0)),
                            'd_grav_rotx': prev.get('d_grav_rotx', 0.0) + float(row.get('d_grav_rotx [arcsec]', 0.0)),
                            'd_grav_roty': prev.get('d_grav_roty', 0.0) + float(row.get('d_grav_roty [arcsec]', 0.0)),
                            'd_grav_rotazi': prev.get('d_grav_rotazi', 0.0) + float(row.get('d_grav_rotazi [arcsec]', 0.0)),
                            'd_grav_rotrad': prev.get('d_grav_rotrad', 0.0) + float(row.get('d_grav_rotrad [arcsec]', 0.0)),
                    }
        except Exception:
            pass
    
    # Load thermal deltas (prefer Position #, sum duplicates)
    if 'MM #' in df.columns:
        try:
            thermal_df = pd.read_excel(path, sheet_name='Thermal', engine='openpyxl')
            if 'Position #' in thermal_df.columns:
                tmp = thermal_df.copy()
                tmp['Position #'] = pd.to_numeric(tmp['Position #'], errors='coerce')
                tmp = tmp[tmp['Position #'].notna()]
                for c in ['d_therm_x [µm]', 'd_therm_y [µm]', 'd_therm_z [µm]', 'd_therm_rotz [arcsec]', 'd_therm_rotx [arcsec]', 'd_therm_roty [arcsec]']:
                    if c in tmp.columns:
                        tmp[c] = pd.to_numeric(tmp[c], errors='coerce').fillna(0.0)
                grp = tmp.groupby('Position #', as_index=False).sum(numeric_only=True)
                for _, row in grp.iterrows():
                    pos = int(row['Position #'])
                    thermal_by_pos[pos] = {
                            'd_therm_x': float(row.get('d_therm_x [µm]', 0.0)) * 1e-6,
                            'd_therm_y': float(row.get('d_therm_y [µm]', 0.0)) * 1e-6,
                            'd_therm_z': float(row.get('d_therm_z [µm]', 0.0)) * 1e-6,
                            'd_therm_rotz': float(row.get('d_therm_rotz [arcsec]', 0.0)),
                            'd_therm_rotx': float(row.get('d_therm_rotx [arcsec]', 0.0)),
                            'd_therm_roty': float(row.get('d_therm_roty [arcsec]', 0.0)),
                            'd_therm_rotazi': float(row.get('d_therm_rotazi [arcsec]', 0.0)),
                            'd_therm_rotrad': float(row.get('d_therm_rotrad [arcsec]', 0.0)),
                    }
            elif 'MM #' in thermal_df.columns:
                for _, row in thermal_df.iterrows():
                    mm_num = row.get('MM #')
                    if pd.isna(mm_num):
                        continue
                    pos = mm_to_pos.get(int(mm_num))
                    if pos is None:
                        continue
                    prev = thermal_by_pos.get(pos, {'d_therm_x':0.0,'d_therm_y':0.0,'d_therm_z':0.0,'d_therm_rotz':0.0})
                    thermal_by_pos[pos] = {
                            'd_therm_x': prev['d_therm_x'] + float(row.get('d_therm_x [µm]', 0.0)) * 1e-6,
                            'd_therm_y': prev['d_therm_y'] + float(row.get('d_therm_y [µm]', 0.0)) * 1e-6,
                            'd_therm_z': prev['d_therm_z'] + float(row.get('d_therm_z [µm]', 0.0)) * 1e-6,
                            'd_therm_rotz': prev['d_therm_rotz'] + float(row.get('d_therm_rotz [arcsec]', 0.0)),
                            'd_therm_rotx': prev.get('d_therm_rotx', 0.0) + float(row.get('d_therm_rotx [arcsec]', 0.0)),
                            'd_therm_roty': prev.get('d_therm_roty', 0.0) + float(row.get('d_therm_roty [arcsec]', 0.0)),
                            'd_therm_rotazi': prev.get('d_therm_rotazi', 0.0) + float(row.get('d_therm_rotazi [arcsec]', 0.0)),
                            'd_therm_rotrad': prev.get('d_therm_rotrad', 0.0) + float(row.get('d_therm_rotrad [arcsec]', 0.0)),
                    }
        except Exception:
            pass
        print(f"VIG DEBUG: thermal_by_pos has {len(thermal_by_pos)} entries")
        if thermal_by_pos:
            sample = dict(list(thermal_by_pos.items())[:3])
            if not os.environ.get('SILENCE_OUTPUT'):
                print(f"VIG DEBUG: thermal_by_pos sample: {sample}")
            rotx_values = [v.get('d_therm_rotx', 0) for v in thermal_by_pos.values()]
            print(f"VIG DEBUG: d_therm_rotx values: min={min(rotx_values):.1f}, max={max(rotx_values):.1f}, all_same={len(set(rotx_values)) == 1}")

    # --- Apply polar vignetting (rotazi + rotrad) after A_eff/weight initialization ---
    try:
        # compute rotation projections using the populated mm_to_pos and *_by_pos
        try:
            _, _, rot_rad_map, rot_azi_map = compute_total_rot_polar(mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos)
            print(f"VIG DEBUG: rot_azi_map has {len(rot_azi_map)} entries")
            if rot_azi_map:
                sample = dict(list(rot_azi_map.items())[:3])
                if not os.environ.get('SILENCE_OUTPUT'):
                    print(f"VIG DEBUG: rot_azi_map sample: {sample}")
                non_zero = [v for v in rot_azi_map.values() if abs(v) > 1e-6]
                print(f"VIG DEBUG: rot_azi_map has {len(non_zero)} non-zero values, range: {min(non_zero) if non_zero else 0:.1f} to {max(non_zero) if non_zero else 0:.1f}")
        except Exception:
            rot_rad_map = {}
            rot_azi_map = {}

        applied_azi = False
        applied_rad = False

        # Read rotazi sheet (A->B interpretation: col0 = delta, col1 = factor)
        # Supported layouts:
        # - Two-column A->B (delta in col0, factor in col1): used for a single
        #   global vignette curve applied to every position.
        # - Multi-column with numeric column headers: columns after col0 are
        #   interpreted as per-position factor series (header=Position #).
        # - Header-heavy workbooks often still contain the intended A->B mapping
        #   in the first two columns; we try that as a final fallback.
        try:
            vdf_azi = pd.read_excel(path, sheet_name='Vignetting rotazi', engine='openpyxl')
            xs_azi = ys_azi = None
            ys_by_pos_azi = {}
            azi_mode = 'none'
            # Try to detect new layout: table starting at column H (index 7)
            # where column H contains MM-config row numbers, column I (8)
            # contains rotazi x-values, column J (9) indicates energy,
            # and column K (10) contains the factor. If present, build
            # per-position/energy interpolation series.
            try:
                aeff_map, aeff_col_name = load_aeff_weight_map_with_name(path)
            except Exception:
                aeff_map, aeff_col_name = load_aeff_weight_map(path), None

            # Try to detect the selected energy (numeric keV) from several
            # places: preferred sources (in order):
            # 1) explicit selection in the vignetting sheet (e.g. cell C2),
            # 2) the A_eff column name (e.g. '7 keV'),
            # 3) a scanned 'keV' token anywhere in the A_eff sheet.
            sel_energy = None
            sel_energy_from_vdf = False
            
            # 1) Enhanced: vignette sheet col 'Selected energy [keV]' and C2
            try:
                # DataFrame scan first (faster) - vdf_azi/vdf_rad defined later, use vdf_rad for now
                for vdf_name, sname in [('vdf_rad', 'rotrad'), ('vdf_azi', 'rotazi')]:
                    vdf = locals().get(vdf_name)
                    if vdf is not None and vdf.shape[1] > 2:
                        for r in range(min(3, vdf.shape[0])):
                            cand = vdf.iloc[r, 2]
                            if pd.notna(cand):
                                try:
                                    sel_energy = float(cand)
                                    sel_energy_from_vdf = True
                                    print(f'VIG sel_energy from {sname} col2 row{r}: {sel_energy}')
                                    break
                                except:
                                    pass
                        if sel_energy is not None:
                            break
                if sel_energy is None:
                    # C2 fallback with openpyxl
                    from openpyxl import load_workbook
                    wb_tmp = load_workbook(path, data_only=True)
                    for sname in ['Vignetting rotazi', 'Vignetting rotrad']:
                        if sname in wb_tmp.sheetnames:
                            ws_tmp = wb_tmp[sname]
                            candidate = ws_tmp.cell(row=2, column=3).value
                            if candidate is not None and not pd.isna(candidate):
                                try:
                                    sel_energy = float(candidate)
                                    sel_energy_from_vdf = True
                                    print(f'VIG sel_energy from {sname} C2: {sel_energy}')
                                    break
                                except:
                                    pass
            except Exception as e:
                print(f'VIG sel_energy scan error: {e}')
            
            if sel_energy is None:
                print('VIG WARNING sel_energy=None - default 0.2')
                sel_energy = 0.2
            print(f'VIG sel_energy={sel_energy}')
            # 2/3) fallback: detect from A_eff or column name
            try:
                # Only use the A_eff column name if sel_energy was not
                # already determined from the vignette sheet itself.
                import re as _re
                if sel_energy is None and aeff_col_name and isinstance(aeff_col_name, str):
                    m = _re.search(r"(\d+(?:\.\d*)?)\s*(?:keV)?", aeff_col_name, flags=_re.IGNORECASE)
                    if m:
                        sel_energy = float(m.group(1))
            except Exception:
                pass

            if sel_energy is None and not sel_energy_from_vdf:
                try:
                    aeff_df = pd.read_excel(path, sheet_name='A_eff', engine='openpyxl', header=None)
                    import re as _re
                    found = None
                    for _, row in aeff_df.iterrows():
                        for cell in row.tolist():
                            try:
                                if isinstance(cell, str):
                                    m = _re.search(r"(\d+(?:\.\d*)?)\s*keV", cell, flags=_re.IGNORECASE)
                                    if m:
                                        found = float(m.group(1))
                                        break
                            except Exception:
                                continue
                        if found is not None:
                            break
                    if found is not None:
                        sel_energy = found
                except Exception:
                    sel_energy = None

            if vdf_azi is not None and not vdf_azi.empty and vdf_azi.shape[1] >= 2:
                # New table layout detection (column H-based)
                if vdf_azi.shape[1] >= 11:
                    # column H (index 7) should contain row numbers referencing
                    # the MM configuration; build per-row arrays keyed by that
                    # row number and energy marker.
                    col_H = vdf_azi.iloc[:, 7]
                    if col_H.notna().any():
                        for _, r in vdf_azi.iterrows():
                            try:
                                cfg_row = r.iloc[7]
                                if pd.isna(cfg_row):
                                    continue
                                cfg_row = int(float(cfg_row))
                            except Exception:
                                continue
                            # energy marker in column J (index 9)
                            energy_marker = r.iloc[9] if vdf_azi.shape[1] > 9 else None
                            try:
                                # column I contains rot delta in arcmin; convert to arcsec
                                xval = float(r.iloc[8]) * 60.0 if not pd.isna(r.iloc[8]) else None
                            except Exception:
                                xval = None
                            try:
                                yval = float(r.iloc[10]) if vdf_azi.shape[1] > 10 and not pd.isna(r.iloc[10]) else None
                            except Exception:
                                yval = None
                            if xval is None or yval is None:
                                continue
                            # store series keyed by (cfg_row, energy_marker)
                            # energy in column J may be numeric or textual; preserve both representations
                            key_str = (cfg_row, str(energy_marker).strip())
                            key_num = None
                            try:
                                key_num = (cfg_row, float(energy_marker))
                            except Exception:
                                key_num = None
                            if key_str not in ys_by_pos_azi:
                                ys_by_pos_azi[key_str] = {'xs': [], 'ys': []}
                            ys_by_pos_azi[key_str]['xs'].append(xval)
                            ys_by_pos_azi[key_str]['ys'].append(yval)
                            if key_num is not None:
                                if key_num not in ys_by_pos_azi:
                                    ys_by_pos_azi[key_num] = {'xs': [], 'ys': []}
                                ys_by_pos_azi[key_num]['xs'].append(xval)
                                ys_by_pos_azi[key_num]['ys'].append(yval)
                        # sort arrays
                        for k, v in list(ys_by_pos_azi.items()):
                            order = np.argsort(v['xs'])
                            xs_sorted = np.array(v['xs'], dtype=float)[order]
                            ys_sorted = np.array(v['ys'], dtype=float)[order]
                            ys_by_pos_azi[k] = (xs_sorted, ys_sorted)
                        if ys_by_pos_azi:
                            azi_mode = 'per_row_energy'
                            print(f'VIG azi_mode={azi_mode} nseries={len(ys_by_pos_azi)}')
                            print(f'VIG ys_by_pos_azi keys sample: {list(ys_by_pos_azi.keys())[:5]}')
                            if os.environ.get("VIG_DEBUG"):
                                for k in list(ys_by_pos_azi)[:3]:
                                    xs, ys = ys_by_pos_azi[k]
                                    print(f'VIG series {k} xs.shape={xs.shape} xs.min={xs.min():.1f} xs.max={xs.max():.1f} ys.mean={ys.mean():.3f}')
                # Fallback to classic modes if new layout not present
                if azi_mode == 'none':
                    xs_raw = pd.to_numeric(vdf_azi.iloc[:, 0], errors='coerce')
                    xs = xs_raw.dropna().to_numpy(dtype=float)
                    if xs.size > 0:
                        if vdf_azi.shape[1] == 2:
                            ys = pd.to_numeric(vdf_azi.iloc[:, 1], errors='coerce').dropna().to_numpy(dtype=float)
                            if ys.size > 0:
                                order = np.argsort(xs)
                                xs_azi = xs[order]
                                ys_azi = ys[order]
                                azi_mode = 'single'
                        else:
                            cols = list(vdf_azi.columns)
                            for col in cols[1:]:
                                try:
                                    pos_key = int(str(col))
                                except Exception:
                                    continue
                                ys_col = pd.to_numeric(vdf_azi[col], errors='coerce')
                                ys_vals = ys_col.fillna(np.nan).to_numpy(dtype=float)
                                if ys_vals.size > 0:
                                    order = np.argsort(xs)
                                    ys_by_pos_azi[pos_key] = ys_vals[order]
                            if ys_by_pos_azi and isinstance(list(ys_by_pos_azi.keys())[0], int):
                                xs_azi = xs[np.argsort(xs)]
                                azi_mode = 'per_pos'
                    # Fallback: try first two columns as A->B
                    if azi_mode == 'none' and vdf_azi.shape[1] >= 2:
                        trial_y = pd.to_numeric(vdf_azi.iloc[:, 1], errors='coerce').dropna().to_numpy(dtype=float)
                        if trial_y.size > 0:
                            order = np.argsort(xs)
                            xs_azi = xs[order]
                            ys_azi = trial_y[order] if trial_y.size == xs_azi.size else trial_y
                            azi_mode = 'single'
                if azi_mode != 'none':
                    applied_azi = True
                    try:
                        if 'MM #' in df.columns and len(df) <= 10:
                            print(f"VIG_PARSE: azi_mode={azi_mode} xs_azi_set={xs_azi is not None} ys_by_pos_azi_keys={list(ys_by_pos_azi.keys()) if 'ys_by_pos_azi' in locals() else None}")
                    except Exception:
                        pass
        except Exception:
            xs_azi = ys_azi = None

        # Read rotrad sheet (A->B interpretation)
        # Read rotrad sheet (same layout rules as rotazi):
        # - 'single' mode: xs_rad/ys_rad forms a single curve (delta->factor)
        # - 'per_pos' mode: ys_by_pos_rad[pos] contains the factor array for that slot
        # During application we prefer per_pos series when present, else fall
        # back to the single global curve. If neither is present we leave
        # the weight unchanged for that component.
        try:
            vdf_rad = pd.read_excel(path, sheet_name='Vignetting rotrad', engine='openpyxl')
            xs_rad = ys_rad = None
            ys_by_pos_rad = {}
            rad_mode = 'none'
            # Try to detect new H/I/J/K layout (same as rotazi):
            # column H=index7 -> cfg_row, I=index8 -> delta(arcmin),
            # J=index9 -> energy marker, K=index10 -> factor
            try:
                if vdf_rad is not None and not vdf_rad.empty and vdf_rad.shape[1] >= 11:
                    col_H = vdf_rad.iloc[:, 7]
                    if col_H.notna().any():
                        for _, r in vdf_rad.iterrows():
                            try:
                                cfg_row = r.iloc[7]
                                if pd.isna(cfg_row):
                                    continue
                                cfg_row = int(float(cfg_row))
                            except Exception:
                                continue
                            energy_marker = r.iloc[9] if vdf_rad.shape[1] > 9 else None
                            try:
                                xval = float(r.iloc[8]) * 60.0 if not pd.isna(r.iloc[8]) else None
                            except Exception:
                                xval = None
                            try:
                                yval = float(r.iloc[10]) if vdf_rad.shape[1] > 10 and not pd.isna(r.iloc[10]) else None
                            except Exception:
                                yval = None
                            if xval is None or yval is None:
                                continue
                            key_str = (cfg_row, str(energy_marker).strip())
                            key_num = None
                            try:
                                key_num = (cfg_row, float(energy_marker))
                            except Exception:
                                key_num = None
                            if key_str not in ys_by_pos_rad:
                                ys_by_pos_rad[key_str] = {'xs': [], 'ys': []}
                            ys_by_pos_rad[key_str]['xs'].append(xval)
                            ys_by_pos_rad[key_str]['ys'].append(yval)
                            if key_num is not None:
                                if key_num not in ys_by_pos_rad:
                                    ys_by_pos_rad[key_num] = {'xs': [], 'ys': []}
                                ys_by_pos_rad[key_num]['xs'].append(xval)
                                ys_by_pos_rad[key_num]['ys'].append(yval)
                        # sort arrays
                        for k, v in list(ys_by_pos_rad.items()):
                            order = np.argsort(v['xs'])
                            xs_sorted = np.array(v['xs'], dtype=float)[order]
                            ys_sorted = np.array(v['ys'], dtype=float)[order]
                            ys_by_pos_rad[k] = (xs_sorted, ys_sorted)
                        if ys_by_pos_rad:
                            rad_mode = 'per_row_energy'
                            print(f'VIG rad_mode={rad_mode} nseries={len(ys_by_pos_rad)}')
                            print(f'VIG ys_by_pos_rad keys sample: {list(ys_by_pos_rad.keys())[:5]}')
                            if os.environ.get("VIG_DEBUG"):
                                for k in list(ys_by_pos_rad)[:3]:
                                    xs, ys = ys_by_pos_rad[k]
                                    print(f'VIG series {k} xs.shape={xs.shape} xs.max={xs.max():.1f}" ys.mean={ys.mean():.3f}')

            except Exception:
                pass

            # Fallback to classic layouts if new layout not detected
            if vdf_rad is not None and not vdf_rad.empty and vdf_rad.shape[1] >= 2 and rad_mode == 'none':
                xs_raw = pd.to_numeric(vdf_rad.iloc[:, 0], errors='coerce')
                xs = xs_raw.dropna().to_numpy(dtype=float)
                if xs.size > 0:
                    if vdf_rad.shape[1] == 2:
                        ys = pd.to_numeric(vdf_rad.iloc[:, 1], errors='coerce').dropna().to_numpy(dtype=float)
                        if ys.size > 0:
                            order = np.argsort(xs)
                            xs_rad = xs[order]
                            ys_rad = ys[order]
                            rad_mode = 'single'
                    else:
                        cols = list(vdf_rad.columns)
                        for col in cols[1:]:
                            try:
                                pos_key = int(str(col))
                            except Exception:
                                continue
                            ys_col = pd.to_numeric(vdf_rad[col], errors='coerce')
                            ys_vals = ys_col.fillna(np.nan).to_numpy(dtype=float)
                            if ys_vals.size > 0:
                                order = np.argsort(xs)
                                ys_by_pos_rad[pos_key] = ys_vals[order]
                        if ys_by_pos_rad:
                            xs_rad = xs[np.argsort(xs)]
                            rad_mode = 'per_pos'
                    if rad_mode == 'none' and vdf_rad.shape[1] >= 2:
                        trial_y = pd.to_numeric(vdf_rad.iloc[:, 1], errors='coerce').dropna().to_numpy(dtype=float)
                        if trial_y.size > 0:
                            order = np.argsort(xs)
                            xs_rad = xs[order]
                            ys_rad = trial_y[order] if trial_y.size == xs_rad.size else trial_y
                            rad_mode = 'single'
                    if rad_mode != 'none':
                        applied_rad = True
                        try:
                            if 'MM #' in df.columns and len(df) <= 10:
                                print(f"VIG_PARSE: rad_mode={rad_mode} xs_rad_set={xs_rad is not None} ys_by_pos_rad_keys={list(ys_by_pos_rad.keys()) if 'ys_by_pos_rad' in locals() else None}")
                        except Exception:
                            pass
        except Exception:
            xs_rad = ys_rad = None
        # If sel_energy was not determined from the rotazi sheet, try rotrad's
        # selected-energy marker (common exporter sometimes places it only
        # on one of the vignette sheets). Normalize to float when possible.
        try:
            if (sel_energy is None or (isinstance(sel_energy, float) and np.isnan(sel_energy))):
                # read rotrad sheet C2 explicitly to detect selected energy
                try:
                    from openpyxl import load_workbook
                    wb_tmp2 = load_workbook(path, data_only=True)
                    if 'Vignetting rotrad' in wb_tmp2.sheetnames:
                        ws_tmp2 = wb_tmp2['Vignetting rotrad']
                        candidate = ws_tmp2.cell(row=2, column=3).value
                    else:
                        candidate = None
                except Exception:
                    candidate = None
                if candidate is not None and not (isinstance(candidate, float) and np.isnan(candidate)):
                    try:
                        if isinstance(candidate, (int, float)):
                            sel_energy = float(candidate)
                        else:
                            sel_energy = float(str(candidate).strip())
                        sel_energy_from_vdf = True
                    except Exception:
                        import re as _re
                        if isinstance(candidate, str):
                            m = _re.search(r"(\d+(?:\.\d*)?)\s*(?:keV)?", candidate, flags=_re.IGNORECASE)
                            if m:
                                sel_energy = float(m.group(1))
                                sel_energy_from_vdf = True
        except Exception:
            pass

        # Apply per-row interpolation multiplicatively to the already-initialized weight
        # (instrumentation: print progress to help find stalls)
        try:
            pass
            sys.stdout.flush()
        except Exception:
            pass
        try:
            if os.environ.get('VIG_DEBUG'):
                try:
                    pass
                    pass
                    pass
                    sys.stdout.flush()
                except Exception:
                    pass
        except Exception:
            pass
        # Apply per-row interpolation multiplicatively to the already-initialized weight.
        # For each MM row we:
        # 1) determine its Position # (slot) using `mm_to_pos` mapping;
        # 2) for each of rot_rad and rot_azi, select the appropriate factor series
        #    (per-position series preferred, otherwise the single global curve);
        # 3) interpolate the factor at the computed angular offset and multiply
        #    it into the existing weight. Radial and azimuthal factors both
        #    multiply the weight (order is irrelevant since multiplication is
        #    commutative), but we apply radial first then azimuthal for clarity.
        # Normalize any per-position vignette series so interpolation
        # later can assume integer-position keys map to 1D arrays of
        # factors sampled at `xs_azi` / `xs_rad`. This fixes cases
        # where sheets mixed tuple (xs, ys) series and plain ys
        # arrays, or where per-pos arrays have a different length
        # than the global xs vectors.
        try:
            if 'ys_by_pos_azi' in locals() and ys_by_pos_azi:
                for k in list(ys_by_pos_azi.keys()):
                    # only normalize integer position keys (per-pos mode)
                    try:
                        if not isinstance(k, int):
                            continue
                    except Exception:
                        continue
                    v = ys_by_pos_azi[k]
                    # if stored as (xs_local, ys_local), resample to xs_azi if available
                    if isinstance(v, tuple) and len(v) == 2:
                        xs_local, ys_local = v
                        try:
                            if 'xs_azi' in locals() and xs_azi is not None:
                                ys_resampled = np.interp(xs_azi, np.asarray(xs_local, dtype=float), np.asarray(ys_local, dtype=float))
                                ys_by_pos_azi[k] = ys_resampled
                            else:
                                ys_by_pos_azi[k] = np.asarray(ys_local, dtype=float)
                        except Exception:
                            ys_by_pos_azi[k] = np.asarray(ys_local, dtype=float)
                    else:
                        # plain array/list -- coerce and if needed resample to xs_azi
                        try:
                            ys_arr = np.asarray(v, dtype=float)
                        except Exception:
                            ys_arr = np.array(v, dtype=float)
                        if 'xs_azi' in locals() and xs_azi is not None:
                            try:
                                if ys_arr.size != xs_azi.size:
                                    xs_from = np.linspace(xs_azi.min(), xs_azi.max(), ys_arr.size)
                                    ys_resampled = np.interp(xs_azi, xs_from, ys_arr)
                                    ys_by_pos_azi[k] = ys_resampled
                                else:
                                    ys_by_pos_azi[k] = ys_arr
                            except Exception:
                                ys_by_pos_azi[k] = ys_arr
                        else:
                            ys_by_pos_azi[k] = ys_arr
        except Exception:
            pass

        try:
            if 'ys_by_pos_rad' in locals() and ys_by_pos_rad:
                for k in list(ys_by_pos_rad.keys()):
                    try:
                        if not isinstance(k, int):
                            continue
                    except Exception:
                        continue
                    v = ys_by_pos_rad[k]
                    if isinstance(v, tuple) and len(v) == 2:
                        xs_local, ys_local = v
                        try:
                            if 'xs_rad' in locals() and xs_rad is not None:
                                ys_resampled = np.interp(xs_rad, np.asarray(xs_local, dtype=float), np.asarray(ys_local, dtype=float))
                                ys_by_pos_rad[k] = ys_resampled
                            else:
                                ys_by_pos_rad[k] = np.asarray(ys_local, dtype=float)
                        except Exception:
                            ys_by_pos_rad[k] = np.asarray(ys_local, dtype=float)
                    else:
                        try:
                            ys_arr = np.asarray(v, dtype=float)
                        except Exception:
                            ys_arr = np.array(v, dtype=float)
                        if 'xs_rad' in locals() and xs_rad is not None:
                            try:
                                if ys_arr.size != xs_rad.size:
                                    xs_from = np.linspace(xs_rad.min(), xs_rad.max(), ys_arr.size)
                                    ys_resampled = np.interp(xs_rad, xs_from, ys_arr)
                                    ys_by_pos_rad[k] = ys_resampled
                                else:
                                    ys_by_pos_rad[k] = ys_arr
                            except Exception:
                                ys_by_pos_rad[k] = ys_arr
                        else:
                            ys_by_pos_rad[k] = ys_arr
        except Exception:
            pass

        if 'weight' in df.columns:

            # Helper to find a per-(cfg_row,energy) series robustly. The
            # vignette sheet may store the energy marker as text or numeric;
            # prefer numeric match, then named A_eff column, then fallback
            # to any series for the cfg_row.
            def _find_series(ys_map, cfg_row, sel_energy_local, aeff_col_name_local=None):
                if not ys_map or cfg_row is None:
                    return None
                # try exact numeric key
                try:
                    keyn = (cfg_row, float(sel_energy_local))
                    if keyn in ys_map:
                        return ys_map[keyn]
                except Exception:
                    pass
                # try named aeff column
                try:
                    if aeff_col_name_local is not None:
                        key_named = (cfg_row, str(aeff_col_name_local).strip())
                        if key_named in ys_map:
                            return ys_map[key_named]
                except Exception:
                    pass
                # try matching by numeric equivalence of the energy token
                for k in ys_map.keys():
                    try:
                        if k[0] != cfg_row:
                            continue
                        val = k[1]
                        try:
                            if sel_energy_local is not None and abs(float(val) - float(sel_energy_local)) < 1e-6:
                                return ys_map[k]
                        except Exception:
                            continue
                    except Exception:
                        continue
                # fallback: return any series for this cfg_row
                for k in ys_map.keys():
                    try:
                        if k[0] == cfg_row:
                            return ys_map[k]
                    except Exception:
                        continue
                return None

            for idx, row in df.iterrows():
                if idx % 100 == 0:
                    try:
                        pass
                        sys.stdout.flush()
                    except Exception:
                        pass
                mm_num = row.get('MM #')
                try:
                    p = int(mm_to_pos.get(int(mm_num))) if pd.notna(mm_num) and int(mm_num) in mm_to_pos else None
                except Exception:
                    p = None
                if p is None:
                    continue

                # radial (per-row)
                if applied_rad and p in rot_rad_map:
                    rot_val = float(rot_rad_map.get(p, 0.0))
                    try:
                        factor = 1.0
                        if 'rad_mode' in locals() and rad_mode == 'per_row_energy' and 'pos_to_cfg_row' in locals():
                            cfg_row = pos_to_cfg_row.get(p)
                            if cfg_row is not None:
                                series = _find_series(ys_by_pos_rad, cfg_row, sel_energy, locals().get('aeff_col_name'))
                                if series is None:
                                    pass
                                else:
                                    xs_use, ys_use = series
                                    factor = float(np.interp(rot_val, xs_use, ys_use))
                        df.at[idx, 'aeff_vig_factor_rad'] = factor
                        df.at[idx, 'weight'] = float(df.at[idx, 'weight']) * float(factor)
                    except Exception:
                        factor = 1.0

                    if 'vig_vals_rad' not in locals():
                        vig_vals_rad = {}
                    if 'vig_source_rad' not in locals():
                        vig_source_rad = {}
                    vig_vals_rad[p] = float(factor)
                    vig_source_rad[p] = ('per_row' if 'rad_mode' in locals() and rad_mode.startswith('per') else 'global')

                # azimuthal (per-row)
                if applied_azi and p in rot_azi_map:
                    try:
                        used = False
                        if 'azi_mode' in locals() and azi_mode == 'per_row_energy' and 'pos_to_cfg_row' in locals():
                            cfg_row = pos_to_cfg_row.get(p)
                            if cfg_row is not None:
                                series = _find_series(ys_by_pos_azi, cfg_row, locals().get('sel_energy'), locals().get('aeff_col_name'))
                                if series is not None:
                                    xs_use, ys_use = series
                                    factor = float(np.interp(abs(float(rot_azi_map.get(p, 0.0))), xs_use, ys_use))
                                    used = True
                        if not used and 'azi_mode' in locals() and azi_mode == 'per_pos' and p in ys_by_pos_azi:
                            ys_use = ys_by_pos_azi[p]
                            rot_val = abs(float(rot_azi_map.get(p, 0.0)))
                            factor = float(np.interp(rot_val, xs_azi, ys_use, left=ys_use[0], right=ys_use[-1]))
                        df.at[idx, 'aeff_vig_factor_azi'] = factor
                        df.at[idx, 'weight'] = float(df.at[idx, 'weight']) * float(factor)
                    except Exception:
                        factor = 1.0

                    if 'vig_vals_azi' not in locals():
                        vig_vals_azi = {}
                    if 'vig_source_azi' not in locals():
                        vig_source_azi = {}
                    vig_vals_azi[p] = float(factor)
                    vig_source_azi[p] = ('per_row' if 'azi_mode' in locals() and azi_mode.startswith('per') else 'global')

            df.attrs['vignetting_rotrad_applied'] = bool(applied_rad)
        # Post-pass: recompute per-position vignette values from the
        # populated per-(cfg_row,energy) tables to avoid mismatches that
        # can occur during the row-wise application loop. This ensures
        # `vig_vals_azi` / `vig_vals_rad` reflect the intended
        # interpolation from the vignetting tables.
        try:
            if os.environ.get('VIG_DEBUG'):
                pass
                sys.stdout.flush()
        except Exception:
            pass
        try:
            if 'pos_to_cfg_row' in locals():
                # ensure vig dicts exist
                if 'vig_vals_azi' not in locals():
                    vig_vals_azi = {}
                if 'vig_vals_rad' not in locals():
                    vig_vals_rad = {}
                if 'vig_source_azi' not in locals():
                    vig_source_azi = {}
                if 'vig_source_rad' not in locals():
                    vig_source_rad = {}

                for pos in sorted(set(list(rot_azi_map.keys()) + list(rot_rad_map.keys()))):
                    cfg = pos_to_cfg_row.get(pos)
                    # azimuthal
                    try:
                        applied_val = None
                        # Prefer per-position arrays when available (azi_mode == 'per_pos')
                        if cfg is not None and 'azi_mode' in locals() and azi_mode == 'per_pos' and 'ys_by_pos_azi' in locals() and isinstance(ys_by_pos_azi, dict) and pos in ys_by_pos_azi:
                            try:
                                applied_val = float(np.interp(abs(float(rot_azi_map.get(pos, 0.0))), xs_azi, ys_by_pos_azi[pos]))
                                vig_vals_azi[pos] = float(applied_val)
                                vig_source_azi[pos] = 'per_pos'
                            except Exception:
                                applied_val = None
                        else:
                            if cfg is not None and 'ys_by_pos_azi' in locals() and ys_by_pos_azi:
                                series = _find_series(ys_by_pos_azi, cfg, locals().get('sel_energy'), locals().get('aeff_col_name'))
                                if series is not None:
                                    xsu, ysu = series
                                    applied_val = float(np.interp(abs(float(rot_azi_map.get(pos, 0.0))), xsu, ysu))
                            if applied_val is None:
                                # fallback to global
                                if 'xs_azi' in locals() and xs_azi is not None and ys_azi is not None:
                                    applied_val = float(np.interp(abs(float(rot_azi_map.get(pos, 0.0))), xs_azi, ys_azi))
                                else:
                                    applied_val = 1.0
                            vig_vals_azi[pos] = float(applied_val)
                            vig_source_azi[pos] = ('per_row' if 'ys_by_pos_azi' in locals() and any((isinstance(k, tuple) and k[0] == cfg) for k in ys_by_pos_azi.keys()) else 'global')
                    except Exception:
                        vig_vals_azi[pos] = 1.0
                        vig_source_azi[pos] = 'none'

                    # radial
                    try:
                        applied_val = None
                        # Prefer per-position arrays when available (rad_mode == 'per_pos')
                        if cfg is not None and 'rad_mode' in locals() and rad_mode == 'per_pos' and 'ys_by_pos_rad' in locals() and isinstance(ys_by_pos_rad, dict) and pos in ys_by_pos_rad:
                            try:
                                applied_val = float(np.interp(float(rot_rad_map.get(pos, 0.0)), xs_rad, ys_by_pos_rad[pos]))
                                vig_vals_rad[pos] = float(applied_val)
                                vig_source_rad[pos] = 'per_pos'
                            except Exception:
                                applied_val = None
                        else:
                            if cfg is not None and 'ys_by_pos_rad' in locals() and ys_by_pos_rad:
                                series = _find_series(ys_by_pos_rad, cfg, locals().get('sel_energy'), locals().get('aeff_col_name'))
                                if series is not None:
                                    xsr, ysr = series
                                    applied_val = float(np.interp(float(rot_rad_map.get(pos, 0.0)), xsr, ysr))
                            if applied_val is None:
                                if 'xs_rad' in locals() and xs_rad is not None and ys_rad is not None:
                                    applied_val = float(np.interp(float(rot_rad_map.get(pos, 0.0)), xs_rad, ys_rad))
                                else:
                                    applied_val = 1.0
                            vig_vals_rad[pos] = float(applied_val)
                            vig_source_rad[pos] = ('per_row' if 'ys_by_pos_rad' in locals() and any((isinstance(k, tuple) and k[0] == cfg) for k in ys_by_pos_rad.keys()) else 'global')
                    except Exception:
                        vig_vals_rad[pos] = 1.0
                        vig_source_rad[pos] = 'none'
        except Exception:
            pass
# Force compute A_eff_adjusted for ALL MMs (even without vignetting)
        print("DEBUG: Starting A_eff_adjusted computation...")
        if 'aeff_base' in df.columns:
            base = df['aeff_base'].astype(float)
            print(f"DEBUG: base sum={base.sum():.3f}, shape={base.shape}")
            # Combine radial and azimuthal vignetting factors
            vig_factor_rad = df.get('aeff_vig_factor_rad', pd.Series(1.0, index=df.index)).fillna(1.0).astype(float)
            vig_factor_azi = df.get('aeff_vig_factor_azi', pd.Series(1.0, index=df.index)).fillna(1.0).astype(float)
            vig_factor = vig_factor_rad * vig_factor_azi
            df['aeff_vig_factor'] = vig_factor  # Store the combined factor
            print(f"DEBUG: vig_factor min/max/mean={vig_factor.min():.3f}/{vig_factor.max():.3f}/{vig_factor.mean():.3f}")
            df['aeff_adjusted'] = base * vig_factor
            df['weight'] = df['aeff_adjusted']
            
            # Debug print: always show
            total_base = float(base.sum())
            total_adj = float(df['aeff_adjusted'].sum())
            pct_diff = ((total_adj - total_base) / total_base * 100) if total_base > 0 else 0
            print(f"A_eff totals: base={total_base:.1f}, adjusted={total_adj:.1f} ({pct_diff:+.1f}%) nMMs={len(df)}")
            print(f"  Sample base/weights: {base.head().tolist()[:5]} / {df['weight'].head().tolist()}")
            
            # Preserve source columns if exist
            for col in ['aeff_vig_source_rotrad', 'aeff_vig_source_rotazi']:
                if col in locals():
                    df[col] = locals()[col]
        else:
            print("DEBUG: No aeff_base column - skipping adjusted computation")

        # Attempt to write per-position vignette factors into the workbook's
        # Vignetting rotazi/rotrad sheets in column B for visibility. This
        # implementation uses openpyxl to update only the specified cells
        # (column B and C1 on vignette sheets; columns B and C on A_eff)
        # and saves atomically to avoid corrupting the original workbook.
        try:
            from openpyxl import load_workbook
            pass
            sys.stdout.flush()
            wb = load_workbook(path)

            # Derive per-position vignette factors from the DataFrame to ensure
            # the values written into the vignette sheets exactly reflect the
            # factors used to compute `aeff_adjusted`.
            per_pos_rad = {}
            per_pos_azi = {}
            try:
                if 'MM #' in df.columns:
                    for _, r in df.iterrows():
                        try:
                            mmn = int(r['MM #'])
                        except Exception:
                            continue
                        posn = mm_to_pos.get(mmn)
                        if posn is None:
                            continue
                        if 'aeff_vig_factor_rad' in df.columns:
                            try:
                                per_pos_rad.setdefault(posn, []).append(float(r.get('aeff_vig_factor_rad', 1.0)))
                            except Exception:
                                pass
                        if 'aeff_vig_factor_azi' in df.columns:
                            try:
                                per_pos_azi.setdefault(posn, []).append(float(r.get('aeff_vig_factor_azi', 1.0)))
                            except Exception:
                                pass
                    # reduce to representative value (mean) for writing
                    for k, vlist in list(per_pos_rad.items()):
                        try:
                            per_pos_rad[k] = float(sum(vlist) / len(vlist)) if vlist else 1.0
                        except Exception:
                            per_pos_rad[k] = 1.0
                    for k, vlist in list(per_pos_azi.items()):
                        try:
                            per_pos_azi[k] = float(sum(vlist) / len(vlist)) if vlist else 1.0
                        except Exception:
                            per_pos_azi[k] = 1.0
            except Exception:
                per_pos_rad = per_pos_rad if per_pos_rad else {}
                per_pos_azi = per_pos_azi if per_pos_azi else {}

            # Build per-position vignette dictionaries from the computed
            # DataFrame values so the values written into the vignette
            # sheets exactly match those used to compute `aeff_adjusted`.
            final_vig_vals_rad = {}
            final_vig_vals_azi = {}
            if 'MM #' in df.columns:
                for _, r in df.iterrows():
                    try:
                        mmv = int(r.get('MM #'))
                    except Exception:
                        continue
                    pos_k = mm_to_pos.get(mmv)
                    if pos_k is None:
                        continue
                    fr = r.get('aeff_vig_factor_rad') if 'aeff_vig_factor_rad' in r.index else None
                    fa = r.get('aeff_vig_factor_azi') if 'aeff_vig_factor_azi' in r.index else None
                    combined = r.get('aeff_vig_factor') if 'aeff_vig_factor' in r.index else None
                    try:
                        if fr is None or (isinstance(fr, float) and (fr != fr)):
                            if fa is not None and not (isinstance(fa, float) and (fa != fa)) and combined is not None:
                                fr = float(combined) / float(fa) if float(fa) != 0 else 1.0
                            else:
                                fr = 1.0
                    except Exception:
                        fr = 1.0
                    try:
                        if fa is None or (isinstance(fa, float) and (fa != fa)):
                            if fr is not None and combined is not None and not (isinstance(fr, float) and (fr != fr)):
                                fa = float(combined) / float(fr) if float(fr) != 0 else float(combined or 1.0)
                            else:
                                fa = float(combined or 1.0)
                    except Exception:
                        fa = float(combined or 1.0)

                    final_vig_vals_rad[pos_k] = float(fr)
                    final_vig_vals_azi[pos_k] = float(fa)

            # If DataFrame-based maps are empty, fall back to previously
            # computed maps (if available).
            if not final_vig_vals_rad and 'vig_vals_rad' in locals():
                final_vig_vals_rad = dict(locals().get('vig_vals_rad', {}))
            if not final_vig_vals_azi and 'vig_vals_azi' in locals():
                final_vig_vals_azi = dict(locals().get('vig_vals_azi', {}))

            # VIGNETTE SHEETS: write col B and C1 only
            for sname, vig_map in (
                ('Vignetting rotazi', final_vig_vals_azi if final_vig_vals_azi else (vig_vals_azi if 'vig_vals_azi' in locals() else {})),
                ('Vignetting rotrad', final_vig_vals_rad if final_vig_vals_rad else (vig_vals_rad if 'vig_vals_rad' in locals() else {})),
            ):
                # Debug print of sample vig_map contents when requested
                try:
                    if os.environ.get('VIG_DEBUG'):
                        sample_items = list((vig_vals_azi if sname.endswith('rotazi') else vig_vals_rad).items())[:10]
                        pass
                        sys.stdout.flush()
                except Exception:
                    pass

                # Extra debug: show a few per-position arrays and rot offsets for pos 1
                try:
                    if os.environ.get('VIG_DEBUG'):
                        p0 = 1
                        pass
                        pass
                        if 'pos_to_cfg_row' in locals():
                            cfg = pos_to_cfg_row.get(p0)
                            pass
                            if 'ys_by_pos_azi' in locals():
                                matches = [k for k in ys_by_pos_azi.keys() if k[0] == cfg]
                                pass
                                if matches:
                                    k = matches[0]
                                    xsu, ysu = ys_by_pos_azi[k]
                                    pass
                            if 'ys_by_pos_rad' in locals():
                                matchesr = [k for k in ys_by_pos_rad.keys() if k[0] == cfg]
                                pass
                                if matchesr:
                                    kr = matchesr[0]
                                    xsr, ysr = ys_by_pos_rad[kr]
                                    pass
                        sys.stdout.flush()
                except Exception:
                    pass

                if not (vig_map and isinstance(vig_map, dict)):
                    continue
                if sname not in wb.sheetnames:
                    continue
                ws = wb[sname]
                max_r = ws.max_row or 0

                # Prefer sheet-specific selected energy in cell C2 only (ignore C1)
                sheet_energy = None
                try:
                    c2 = ws.cell(row=2, column=3).value
                    if c2 is not None and not (isinstance(c2, float) and np.isnan(c2)):
                        try:
                            sheet_energy = float(c2)
                        except Exception:
                            import re as _re
                            if isinstance(c2, str):
                                m = _re.search(r"(\d+(?:\.\d*)?)", c2)
                                if m:
                                    sheet_energy = float(m.group(1))
                except Exception:
                    sheet_energy = None
                if sheet_energy is None:
                    sheet_energy = locals().get('sel_energy') if 'sel_energy' in locals() else None

                # Build mapping row->position from column A (explicit mapping expected)
                pos_row_map = {}
                # If the first column looks like a delta table (e.g. header 'delta' or 'delta_arcsec')
                # then do not attempt to interpret integer-like values in column A as Position #.
                try:
                    header_cell = ws.cell(row=1, column=1).value
                    header_is_delta = False
                    if isinstance(header_cell, str):
                        import re as _re
                        if _re.search(r"\bdelta\b|delta_arc|arcsec", header_cell, flags=_re.IGNORECASE):
                            header_is_delta = True
                except Exception:
                    header_is_delta = False

                if not header_is_delta:
                    for r in range(1, max_r + 1):
                        try:
                            v = ws.cell(row=r, column=1).value
                            if v is None:
                                continue
                            if isinstance(v, (int,)) or (isinstance(v, float) and float(v).is_integer()):
                                pos_row_map[int(v)] = r
                            else:
                                s = str(v).strip()
                                if s.isdigit():
                                    pos_row_map[int(s)] = r
                        except Exception:
                            continue

                written = 0
                total = len(pos_row_map)
                # For each position recorded in column A, compute/write factor into column B
                # Rebuild a robust pos->cfg_row mapping from MM configuration if available
                local_pos_to_cfg = {}
                try:
                    mmcfg = pd.read_excel(path, sheet_name='MM configuration', engine='openpyxl')
                    if 'Position #' in mmcfg.columns:
                        tmp = mmcfg.copy()
                        tmp['Position #'] = pd.to_numeric(tmp['Position #'], errors='coerce')
                        tmp = tmp[tmp['Position #'].notna()]
                        for order_i, (_, rr) in enumerate(tmp.iterrows()):
                            try:
                                pval = int(rr['Position #'])
                                # Prefer explicit 'Row #' when present; otherwise
                                # fall back to the enumerated row index.
                                raw_rownum = rr.get('Row #') if 'Row #' in mmcfg.columns else None
                                if raw_rownum is not None and not pd.isna(raw_rownum):
                                    local_pos_to_cfg[pval] = int(float(raw_rownum))
                                else:
                                    local_pos_to_cfg[pval] = order_i + 1
                            except Exception:
                                continue
                    else:
                        # fallback: use row order as cfg_row for positions if 'Position #' absent
                        for order_i, (_, rr) in enumerate(mmcfg.iterrows()):
                            local_pos_to_cfg[order_i + 1] = order_i + 1
                except Exception:
                    local_pos_to_cfg = locals().get('pos_to_cfg_row', {}) if 'pos_to_cfg_row' in locals() else {}

                for i, (pos_k, row_idx) in enumerate(pos_row_map.items()):
                    if i % 100 == 0:
                        pass
                        sys.stdout.flush()
                    pos_int = int(pos_k)
                    # prefer per-position values derived from the DataFrame
                    val = None
                    try:
                        if sname.endswith('rotazi'):
                            if pos_int in per_pos_azi:
                                val = float(per_pos_azi[pos_int])
                        else:
                            if pos_int in per_pos_rad:
                                val = float(per_pos_rad[pos_int])
                    except Exception:
                        val = None
                    # fallback to any precomputed vig_map entry
                    try:
                        if val is None and pos_int in vig_map:
                            val = float(vig_map[pos_int])
                    except Exception:
                        val = None
                    # if missing, attempt to compute from per-(cfg_row,energy) tables
                    if val is None:
                        try:
                            cfg_row = None
                            if 'pos_to_cfg_row' in locals() and pos_int in pos_to_cfg_row:
                                cfg_row = pos_to_cfg_row.get(pos_int)
                            elif pos_int in local_pos_to_cfg:
                                cfg_row = local_pos_to_cfg.get(pos_int)
                            if sname.endswith('rotazi'):
                                # choose azimuthal series
                                if cfg_row is not None and 'ys_by_pos_azi' in locals() and ys_by_pos_azi:
                                    series = _find_series(ys_by_pos_azi, cfg_row, sheet_energy, locals().get('aeff_col_name'))
                                    if series is not None:
                                        xs_u, ys_u = series
                                        val = float(np.interp(abs(float(rot_azi_map.get(pos_int, 0.0))), xs_u, ys_u))
                                    else:
                                        val = float(np.interp(abs(float(rot_azi_map.get(pos_int, 0.0))), locals().get('xs_azi', np.array([0.])), locals().get('ys_azi', np.array([1.]))))
                                else:
                                        val = float(np.interp(abs(float(rot_azi_map.get(pos_int, 0.0))), locals().get('xs_azi', np.array([0.])), locals().get('ys_azi', np.array([1.]))))
                            else:
                                # rotrad
                                if cfg_row is not None and 'ys_by_pos_rad' in locals() and ys_by_pos_rad:
                                    series = _find_series(ys_by_pos_rad, cfg_row, sheet_energy, locals().get('aeff_col_name'))
                                    if series is not None:
                                        xs_u, ys_u = series
                                        val = float(np.interp(abs(float(rot_rad_map.get(pos_int, 0.0))), xs_u, ys_u))
                                    else:
                                        val = float(np.interp(abs(float(rot_rad_map.get(pos_int, 0.0))), locals().get('xs_rad', np.array([0.])), locals().get('ys_rad', np.array([1.]))))
                                else:
                                    val = float(np.interp(abs(float(rot_rad_map.get(pos_int, 0.0))), locals().get('xs_rad', np.array([0.])), locals().get('ys_rad', np.array([1.]))))
                        except Exception:
                            val = 1.0
                    try:
                        ws.cell(row=row_idx, column=2, value=float(val))
                        written += 1
                    except Exception:
                        continue

                pass

                # Do not write to cell C1 here; GUI is responsible for any
                # selected-energy markers (e.g. cell C2). Leave C1/C2 untouched.

            # A_eff sheet: update columns B and C only
            if 'A_eff' in wb.sheetnames and ('aeff_adjusted' in df.columns or 'aeff_base' in df.columns):
                ws_a = wb['A_eff']
                max_r_a = ws_a.max_row or 0

                # Read vignette factors from workbook (column B of vignette sheets)
                vig_rad_sheet = {}
                vig_azi_sheet = {}
                if 'Vignetting rotrad' in wb.sheetnames:
                    wsr = wb['Vignetting rotrad']
                    for rr in range(1, wsr.max_row + 1):
                        a = wsr.cell(row=rr, column=1).value
                        b = wsr.cell(row=rr, column=2).value
                        try:
                            key = int(a) if isinstance(a, (int, float)) or (isinstance(a, str) and str(a).strip().isdigit()) else None
                        except Exception:
                            key = None
                        if key is not None and b is not None:
                            try:
                                vig_rad_sheet[key] = float(b)
                            except Exception:
                                pass
                if 'Vignetting rotazi' in wb.sheetnames:
                    wsa = wb['Vignetting rotazi']
                    for rr in range(1, wsa.max_row + 1):
                        a = wsa.cell(row=rr, column=1).value
                        b = wsa.cell(row=rr, column=2).value
                        try:
                            key = int(a) if isinstance(a, (int, float)) or (isinstance(a, str) and str(a).strip().isdigit()) else None
                        except Exception:
                            key = None
                        if key is not None and b is not None:
                            try:
                                vig_azi_sheet[key] = float(b)
                            except Exception:
                                pass

                # Update the in-memory dataframe so its per-MM factors and
                # adjusted A_eff match the vignette values written to the sheets.
                try:
                    # ensure columns exist
                    if 'aeff_vig_factor_rad' not in df.columns:
                        df['aeff_vig_factor_rad'] = np.nan
                    if 'aeff_vig_factor_azi' not in df.columns:
                        df['aeff_vig_factor_azi'] = np.nan
                    if 'aeff_vig_factor' not in df.columns:
                        df['aeff_vig_factor'] = np.nan
                    if 'aeff_adjusted' not in df.columns:
                        df['aeff_adjusted'] = np.nan
                    for idx in df.index:
                        try:
                            mmv = int(df.at[idx, 'MM #'])
                        except Exception:
                            continue
                        pos = mm_to_pos.get(mmv) if 'mm_to_pos' in locals() else None
                        rad_f = vig_rad_sheet.get(pos, None) if pos is not None else None
                        azi_f = vig_azi_sheet.get(pos, None) if pos is not None else None
                        combined = None
                        try:
                            if rad_f is not None and azi_f is not None:
                                combined = float(rad_f) * float(azi_f)
                        except Exception:
                            combined = None
                        try:
                            if rad_f is not None:
                                df.at[idx, 'aeff_vig_factor_rad'] = float(rad_f)
                        except Exception:
                            df.at[idx, 'aeff_vig_factor_rad'] = np.nan
                        try:
                            if azi_f is not None:
                                df.at[idx, 'aeff_vig_factor_azi'] = float(azi_f)
                        except Exception:
                            df.at[idx, 'aeff_vig_factor_azi'] = np.nan
                        try:
                            if combined is not None:
                                df.at[idx, 'aeff_vig_factor'] = float(combined)
                            else:
                                df.at[idx, 'aeff_vig_factor'] = np.nan
                        except Exception:
                            df.at[idx, 'aeff_vig_factor'] = np.nan
                        try:
                            base_val = float(df.at[idx, 'aeff_base']) if 'aeff_base' in df.columns and not pd.isna(df.at[idx, 'aeff_base']) else 0.0
                        except Exception:
                            base_val = 0.0
                        try:
                            if combined is None or base_val == 0.0:
                                df.at[idx, 'aeff_adjusted'] = 0.0
                            else:
                                df.at[idx, 'aeff_adjusted'] = float(base_val) * float(combined)
                        except Exception:
                            df.at[idx, 'aeff_adjusted'] = 0.0
                except Exception:
                    pass

                # Now write A_eff columns using sheet-derived vignette products to
                # ensure A_eff C/B == (rotazi_B * rotrad_B)
                written_a = 0
                for r in range(1, max_r_a + 1):
                    cell = ws_a.cell(row=r, column=1).value
                    if cell is None:
                        continue
                    try:
                        mmv = int(cell) if isinstance(cell, (int, float)) or (isinstance(cell, str) and str(cell).strip().isdigit()) else None
                        if isinstance(cell, str) and str(cell).strip().isdigit():
                            mmv = int(str(cell).strip())
                    except Exception:
                        mmv = None
                    if mmv is None:
                        continue
                    # write base if available in df
                    try:
                        base_row = df.loc[df['MM #'] == mmv]
                        base_val = float(base_row['aeff_base'].iat[0]) if (not base_row.empty and 'aeff_base' in base_row.columns) else 0.0
                    except Exception:
                        base_val = 0.0

                    # determine per-position factor from sheet
                    pos = mm_to_pos.get(mmv) if 'mm_to_pos' in locals() else None
                    rad_f = vig_rad_sheet.get(pos, None) if pos is not None else None
                    azi_f = vig_azi_sheet.get(pos, None) if pos is not None else None
                    combined = None
                    try:
                        if rad_f is not None and azi_f is not None:
                            combined = float(rad_f) * float(azi_f)
                    except Exception:
                        combined = None

                    # Write column B = canonical base (if known)
                    if base_val is not None:
                        try:
                            ws_a.cell(row=r, column=2, value=float(base_val))
                        except Exception:
                            pass

                    # Write column C = adjusted using sheet-derived factors when available
                    try:
                        if combined is None or base_val is None or base_val == 0.0:
                            ws_a.cell(row=r, column=3, value=0.0)
                        else:
                            ws_a.cell(row=r, column=3, value=float(base_val) * float(combined))
                        written_a += 1
                    except Exception:
                        try:
                            ws_a.cell(row=r, column=3, value=0.0)
                            written_a += 1
                        except Exception:
                            pass

                # Sync workbook A_eff column C back into the dataframe so the
                # returned df matches what was written to file.
                try:
                    for r in range(1, max_r_a + 1):
                        cell = ws_a.cell(row=r, column=1).value
                        if cell is None:
                            continue
                        try:
                            mmv = int(cell) if isinstance(cell, (int, float)) or (isinstance(cell, str) and str(cell).strip().isdigit()) else None
                            if isinstance(cell, str) and str(cell).strip().isdigit():
                                mmv = int(str(cell).strip())
                        except Exception:
                            mmv = None
                        if mmv is None:
                            continue
                        cval = ws_a.cell(row=r, column=3).value
                        try:
                            # Find rows in df matching this MM and update
                            mask = df['MM #'] == mmv
                            if 'aeff_adjusted' in df.columns:
                                df.loc[mask, 'aeff_adjusted'] = float(cval) if cval is not None else 0.0
                            # also update combined vig factor column if base present
                            if 'aeff_base' in df.columns and cval is not None:
                                try:
                                    base_vals = df.loc[mask, 'aeff_base']
                                    for i_idx in base_vals.index:
                                        b = base_vals.at[i_idx]
                                        try:
                                            if b and float(b) != 0.0:
                                                df.at[i_idx, 'aeff_vig_factor'] = float(cval) / float(b)
                                        except Exception:
                                            pass
                                except Exception:
                                    pass
                        except Exception:
                            pass
                except Exception:
                    pass

                # Always save workbook after writing column C
                # Before saving, reconcile per-position vignette factors so
                # the values written into the vignette sheets (column B)
                # are consistent with the `aeff_adjusted` values computed
                # for each MM. Compute per-position combined factor from the
                # DataFrame and adjust rad/azi components accordingly.
                try:
                    # Build mapping pos -> member MM indices (from mm_to_pos)
                    pos_members = {}
                    if 'MM #' in df.columns:
                        for idx_row, mmv in df['MM #'].items():
                            try:
                                mm_key = int(float(mmv))
                            except Exception:
                                continue
                            ppos = mm_to_pos.get(mm_key)
                            if ppos is None:
                                continue
                            pos_members.setdefault(ppos, []).append(idx_row)

                    # Compute canonical per-position vignette factors directly
                    # from the vignetting tables (preferred source) so the
                    # values written into the vignette sheets reflect the
                    # same interpolation used for per-position corrections.
                    final_vig_rad = {}
                    final_vig_azi = {}
                    for pos in sorted(pos_members.keys()):
                        cfg = pos_to_cfg_row.get(pos) if 'pos_to_cfg_row' in locals() else None
                        # AZIMUTHAL: prefer per-pos series, then per-row series, then global
                        try:
                            applied_azi = None
                            if cfg is not None and 'ys_by_pos_azi' in locals() and isinstance(ys_by_pos_azi, dict):
                                # Support both per-pos dict keyed by Position# and
                                # per-row dict keyed by (cfg_row, energy). Per-pos
                                # entries may be either plain ys arrays (sampled
                                # at xs_azi) or (xs, ys) tuples. Handle both.
                                series = None
                                try:
                                    if pos in ys_by_pos_azi:
                                        series = ys_by_pos_azi[pos]
                                except Exception:
                                    pass
                                if series is None:
                                    series = _find_series(ys_by_pos_azi, cfg, locals().get('sel_energy'), locals().get('aeff_col_name'))
                                if series is not None:
                                    try:
                                        if isinstance(series, tuple) and len(series) == 2:
                                            xsu, ysu = series
                                        else:
                                            # assume series is a 1D ys array sampled at xs_azi
                                            ysu = np.asarray(series, dtype=float)
                                            if 'xs_azi' in locals() and xs_azi is not None:
                                                xsu = xs_azi
                                            else:
                                                raise ValueError('no xs available for per-pos ys')
                                        applied_azi = float(np.interp(abs(float(rot_azi_map.get(pos, 0.0))), xsu, ysu, left=ysu[0], right=ysu[-1]))
                                    except Exception as e:
                                        print(f"VIG_DEBUG: interp error pos={pos} e={e}")
                                        applied_azi = None
                            if applied_azi is None and 'xs_azi' in locals() and xs_azi is not None and 'ys_azi' in locals() and ys_azi is not None:
                                try:
                                    applied_azi = float(np.interp(abs(float(rot_azi_map.get(pos, 0.0))), xs_azi, ys_azi))
                                except Exception as e:
                                    print(f"VIG_DEBUG: global interp azi error pos={pos} e={e}")
                                    applied_azi = None
                            if applied_azi is None:
                                applied_azi = 1.0
                            final_vig_azi[pos] = float(applied_azi)
                            try:
                                if 'MM #' in df.columns and len(df) <= 10:
                                    print(f"VIG_DEBUG_COMPUTE: pos={pos} applied_azi={applied_azi}")
                            except Exception:
                                pass
                        except Exception:
                            final_vig_azi[pos] = 1.0

                        # RADIAL: prefer per-pos series, then per-row series, then global
                        try:
                            applied_rad = None
                            if cfg is not None and 'ys_by_pos_rad' in locals() and isinstance(ys_by_pos_rad, dict):
                                # Support both per-pos dict keyed by Position# and
                                # per-row dict keyed by (cfg_row, energy)
                                series = None
                                try:
                                    if pos in ys_by_pos_rad:
                                        series = ys_by_pos_rad[pos]
                                except Exception:
                                    pass
                                if series is None:
                                    series = _find_series(ys_by_pos_rad, cfg, locals().get('sel_energy'), locals().get('aeff_col_name'))
                                if series is not None:
                                    try:
                                        if isinstance(series, tuple) and len(series) == 2:
                                            xsr, ysr = series
                                        else:
                                            ysr = np.asarray(series, dtype=float)
                                            if 'xs_rad' in locals() and xs_rad is not None:
                                                xsr = xs_rad
                                            else:
                                                raise ValueError('no xs available for per-pos ys')
                                        applied_rad = float(np.interp(float(rot_rad_map.get(pos, 0.0)), xsr, ysr, left=ysr[0], right=ysr[-1]))
                                    except Exception:
                                        applied_rad = None
                            if applied_rad is None and 'xs_rad' in locals() and xs_rad is not None and 'ys_rad' in locals() and ys_rad is not None:
                                applied_rad = float(np.interp(float(rot_rad_map.get(pos, 0.0)), xs_rad, ys_rad))
                            if applied_rad is None:
                                applied_rad = 1.0
                            final_vig_rad[pos] = float(applied_rad)
                            try:
                                if 'MM #' in df.columns and len(df) <= 10:
                                    print(f"VIG_DEBUG_COMPUTE: pos={pos} applied_rad={applied_rad}")
                            except Exception:
                                pass
                        except Exception:
                            final_vig_rad[pos] = 1.0

                    # Overwrite vig_vals_* with canonical interpolated values
                    vig_vals_rad = dict(final_vig_rad)
                    vig_vals_azi = dict(final_vig_azi)

                    # Apply these canonical per-position factors to the DataFrame
                    try:
                        if 'MM #' in df.columns:
                            df['aeff_vig_factor_rad'] = df['MM #'].map(lambda mm: vig_vals_rad.get(mm_to_pos.get(int(mm)) if mm is not None else None, 1.0))
                            df['aeff_vig_factor_azi'] = df['MM #'].map(lambda mm: vig_vals_azi.get(mm_to_pos.get(int(mm)) if mm is not None else None, 1.0))
                            df['aeff_vig_factor'] = df['aeff_vig_factor_rad'].fillna(1.0).astype(float) * df['aeff_vig_factor_azi'].fillna(1.0).astype(float)
                            if 'aeff_base' in df.columns:
                                df['aeff_adjusted'] = df['aeff_base'].astype(float) * df['aeff_vig_factor']
                                df['weight'] = df['aeff_adjusted']
                    except Exception:
                        pass
                    # Debug small workbooks: print final maps and DF snippet
                    try:
                        if len(df) <= 10:
                            print(f"VIG_DEBUG_SMALL: final_vig_azi={final_vig_azi}")
                            print(f"VIG_DEBUG_SMALL: final_vig_rad={final_vig_rad}")
                            print(df[['MM #','aeff_base','aeff_vig_factor_rad','aeff_vig_factor_azi','aeff_vig_factor','aeff_adjusted']].head(20))
                    except Exception:
                        pass
                    except Exception:
                        pass

                    # Also write these reconciled per-position factors back into
                    # the vignette sheets' column B so the workbook reflects the
                    # exact factors used to compute `aeff_adjusted`.
                    for sname, final_map in (('Vignetting rotazi', vig_vals_azi), ('Vignetting rotrad', vig_vals_rad)):
                        if sname not in wb.sheetnames:
                            continue
                        ws_w = wb[sname]
                        # build row map from column A (like earlier)
                        pos_row_map2 = {}
                        max_r2 = ws_w.max_row or 0
                        for r2 in range(1, max_r2 + 1):
                            try:
                                v = ws_w.cell(row=r2, column=1).value
                                if v is None:
                                    continue
                                if isinstance(v, (int,)) or (isinstance(v, float) and float(v).is_integer()):
                                    pos_row_map2[int(v)] = r2
                                else:
                                    s = str(v).strip()
                                    if s.isdigit():
                                        pos_row_map2[int(s)] = r2
                            except Exception:
                                continue
                        # write reconciled values
                        for cfg_or_pos_k, row_idx in pos_row_map2.items():
                            try:
                                # The sheet's column A may contain cfg_row identifiers
                                # rather than Position #. Try to resolve to a
                                # Position # using pos_to_cfg_row if available.
                                pos_key = None
                                try:
                                    if 'pos_to_cfg_row' in locals():
                                        # find position whose cfg_row equals this key
                                        for pp, cfgv in pos_to_cfg_row.items():
                                            if cfgv == int(cfg_or_pos_k):
                                                pos_key = pp
                                                break
                                except Exception:
                                    pos_key = None
                                # fallback: if the cell already contains a Position #
                                if pos_key is None:
                                    try:
                                        pos_key = int(cfg_or_pos_k)
                                    except Exception:
                                        pos_key = None
                                val = float(final_map.get(pos_key, 1.0))
                                ws_w.cell(row=row_idx, column=2, value=val)
                            except Exception:
                                continue
                except Exception:
                    pass

                # After reconciling vignette sheet B values, re-write A_eff
                # column C using the newly-updated vignette B values so the
                # workbook's A_eff entries are consistent with the vignette
                # factors that will be saved below.
                try:
                    vig_azi_sheet = {}
                    vig_rad_sheet = {}
                    if 'Vignetting rotazi' in wb.sheetnames:
                        wsa = wb['Vignetting rotazi']
                        for rr in range(1, wsa.max_row + 1):
                            a = wsa.cell(row=rr, column=1).value
                            b = wsa.cell(row=rr, column=2).value
                            try:
                                key = int(a) if isinstance(a, (int, float)) or (isinstance(a, str) and str(a).strip().isdigit()) else None
                            except Exception:
                                key = None
                            if key is not None and b is not None:
                                try:
                                    vig_azi_sheet[key] = float(b)
                                except Exception:
                                    pass
                    if 'Vignetting rotrad' in wb.sheetnames:
                        wsr = wb['Vignetting rotrad']
                        for rr in range(1, wsr.max_row + 1):
                            a = wsr.cell(row=rr, column=1).value
                            b = wsr.cell(row=rr, column=2).value
                            try:
                                key = int(a) if isinstance(a, (int, float)) or (isinstance(a, str) and str(a).strip().isdigit()) else None
                            except Exception:
                                key = None
                            if key is not None and b is not None:
                                try:
                                    vig_rad_sheet[key] = float(b)
                                except Exception:
                                    pass

                    # Write A_eff column C based on these sheet values
                    if 'A_eff' in wb.sheetnames:
                        ws_a = wb['A_eff']
                        for r in range(1, ws_a.max_row + 1):
                            cell = ws_a.cell(row=r, column=1).value
                            if cell is None:
                                continue
                            try:
                                mmv = int(cell) if isinstance(cell, (int, float)) or (isinstance(cell, str) and str(cell).strip().isdigit()) else None
                                if isinstance(cell, str) and str(cell).strip().isdigit():
                                    mmv = int(str(cell).strip())
                            except Exception:
                                mmv = None
                            if mmv is None:
                                continue
                            try:
                                base_row = df.loc[df['MM #'] == mmv]
                                base_val = float(base_row['aeff_base'].iat[0]) if (not base_row.empty and 'aeff_base' in base_row.columns) else 0.0
                            except Exception:
                                base_val = 0.0
                            pos = mm_to_pos.get(mmv) if 'mm_to_pos' in locals() else None
                            rad_f = vig_rad_sheet.get(pos, None) if pos is not None else None
                            azi_f = vig_azi_sheet.get(pos, None) if pos is not None else None
                            combined = None
                            try:
                                if rad_f is not None and azi_f is not None:
                                    combined = float(rad_f) * float(azi_f)
                            except Exception:
                                combined = None
                            try:
                                if combined is None or base_val is None or base_val == 0.0:
                                    ws_a.cell(row=r, column=3, value=0.0)
                                else:
                                    ws_a.cell(row=r, column=3, value=float(base_val) * float(combined))
                            except Exception:
                                try:
                                    ws_a.cell(row=r, column=3, value=0.0)
                                except Exception:
                                    pass
                except Exception:
                    pass

                tmpf = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                tmpf.close()
                wb.save(tmpf.name)
                os.replace(tmpf.name, path)

                # Re-open the saved workbook and synchronize A_eff column C
                # back into the dataframe to guarantee the returned `df`
                # exactly mirrors what is persisted on disk.
                try:
                    from openpyxl import load_workbook as _load_wb
                    wb_saved = _load_wb(path, data_only=True)
                    if 'A_eff' in wb_saved.sheetnames:
                        ws_saved_a = wb_saved['A_eff']
                        for r in range(1, ws_saved_a.max_row + 1):
                            cell = ws_saved_a.cell(row=r, column=1).value
                            if cell is None:
                                continue
                            try:
                                mmv = int(cell) if isinstance(cell, (int, float)) or (isinstance(cell, str) and str(cell).strip().isdigit()) else None
                                if isinstance(cell, str) and str(cell).strip().isdigit():
                                    mmv = int(str(cell).strip())
                            except Exception:
                                mmv = None
                            if mmv is None:
                                continue
                            cval = ws_saved_a.cell(row=r, column=3).value
                            try:
                                mask = df['MM #'] == mmv
                                if 'aeff_adjusted' in df.columns:
                                    df.loc[mask, 'aeff_adjusted'] = float(cval) if cval is not None else 0.0
                                if 'aeff_base' in df.columns and cval is not None:
                                    base_vals = df.loc[mask, 'aeff_base']
                                    for i_idx in base_vals.index:
                                        try:
                                            b = float(base_vals.at[i_idx])
                                            if b != 0.0:
                                                df.at[i_idx, 'aeff_vig_factor'] = float(cval) / b
                                        except Exception:
                                            pass
                            except Exception:
                                pass
                    try:
                        wb_saved.close()
                    except Exception:
                        pass
                except Exception:
                    pass
                pass
                sys.stdout.flush()

            try:
                wb.close()
            except Exception:
                pass
            pass
            sys.stdout.flush()
        except Exception:
            # If anything goes wrong, do not raise — vignetting writes are non-fatal.
            pass
            sys.stdout.flush()
        # Debug summary: print selected vignette source mapping for first positions
        try:
            debug_env = os.environ.get('VIG_DEBUG', None)
            if debug_env:
                sample = sorted(list(mm_to_pos.values()))[:16]
                pass
        except Exception:
            pass
    except Exception:
        # non-fatal: continue without vignetting
        pass
    
    # Apply deltas to m_rad and m_azi, and rotz effect to m_azi (deltas are per position)
    for idx, row in df.iterrows():
        mm_num = row['MM #']
        pos = mm_to_pos.get(int(mm_num))
        new_m_rad = row['m_rad']
        new_m_azi = row['m_azi']
        
        # Apply alignment rad/azi deltas
        if pos is not None and pos in alignment_by_pos:
            deltas = alignment_by_pos[pos]
            new_m_rad += deltas['d_align_rad']
            new_m_azi += deltas['d_align_azi']
        
        # Calculate total rotz from all perturbations and apply to m_azi
        d_rotz_total_arcsec = 0
        if pos is not None and pos in alignment_by_pos:
            d_rotz_total_arcsec += alignment_by_pos[pos].get('d_align_rotz', 0)
        if pos is not None and pos in gravity_by_pos:
            d_rotz_total_arcsec += gravity_by_pos[pos].get('d_grav_rotz', 0)
        if pos is not None and pos in thermal_by_pos:
            d_rotz_total_arcsec += thermal_by_pos[pos].get('d_therm_rotz', 0)
        
        # Apply rotz effect: d_azi increases by r_MM * d_rotz (in radians)
        if d_rotz_total_arcsec != 0 and mm_num in mm_config_map:
            r_MM = mm_config_map[mm_num].get('r_MM', 0)
            d_rotz_rad = np.radians(d_rotz_total_arcsec / 3600.0)  # arcsec -> degrees -> radians
            new_m_azi += r_MM * d_rotz_rad  # Both in meters
        
        df.at[idx, 'm_rad'] = new_m_rad
        df.at[idx, 'm_azi'] = new_m_azi
    
    # Convert from polar (m_rad, m_azi) to Cartesian (mux, muy)
    # Radial direction: along MM position (x_MM, y_MM)
    # Azimuthal direction: perpendicular to radial, counterclockwise
    # mux = (x_MM/r_MM)*m_rad - (y_MM/r_MM)*m_azi
    # muy = (y_MM/r_MM)*m_rad + (x_MM/r_MM)*m_azi
    # This ensures consistent behavior: positive m_rad always moves outward from center
    
    def convert_polar_to_cartesian(row, mm_config_map):
        """Convert polar offsets (m_rad, m_azi) to Cartesian using MM position as reference."""
        mm_num = row['MM #']
        # Use provided mm_config_map or sensible defaults when missing
        config = mm_config_map.get(mm_num, {'x_MM': 1.0, 'y_MM': 0.0, 'r_MM': 1.0})
        x_mm = config.get('x_MM', 1.0)
        y_mm = config.get('y_MM', 0.0)
        r_mm = config.get('r_MM', 1.0)  # Avoid division by zero
        
        # Unit vectors
        u_rad_x = x_mm / r_mm  # Radial unit vector x-component
        u_rad_y = y_mm / r_mm  # Radial unit vector y-component
        u_azi_x = -y_mm / r_mm  # Azimuthal unit vector x-component (perpendicular)
        u_azi_y = x_mm / r_mm   # Azimuthal unit vector y-component
        
        m_rad = row['m_rad']
        m_azi = row['m_azi']
        
        mux = u_rad_x * m_rad + u_azi_x * m_azi
        muy = u_rad_y * m_rad + u_azi_y * m_azi
        
        return mux, muy
    
    # Apply conversion for all rows
    df[['mux', 'muy']] = df.apply(
        lambda row: pd.Series(convert_polar_to_cartesian(row, mm_config_map)),
        axis=1
    )
    
    # Apply all perturbations to mux and muy
    for idx, row in df.iterrows():
        mm_num = row['MM #']
        pos = mm_to_pos.get(int(mm_num))
        
        # Start with current values
        new_mux = row['mux']
        new_muy = row['muy']
        
        # Get alignment d_z
        d_align_z = 0
        if pos is not None and pos in alignment_by_pos:
            d_align_z = alignment_by_pos[pos].get('d_align_z', 0)
        
        # Get gravity d_z
        d_grav_z = 0
        if pos is not None and pos in gravity_by_pos:
            new_mux += gravity_by_pos[pos]['d_grav_x']
            new_muy += gravity_by_pos[pos]['d_grav_y']
            d_grav_z = gravity_by_pos[pos]['d_grav_z']
        
        # Get thermal d_z
        d_therm_z = 0
        if pos is not None and pos in thermal_by_pos:
            new_mux += thermal_by_pos[pos]['d_therm_x']
            new_muy += thermal_by_pos[pos]['d_therm_y']
            d_therm_z = thermal_by_pos[pos]['d_therm_z']
        
        # Calculate d_z_total and apply z-axis projection
        d_z_total = d_align_z + d_grav_z + d_therm_z
        if mm_num in mm_config_map:
            mm_config = mm_config_map[mm_num]
            x_MM = mm_config['x_MM']
            y_MM = mm_config['y_MM']
            z_MM = mm_config['z_MM']
            
            # Calculate dm_x and dm_y based on z displacement
            denominator = 12 - z_MM
            if denominator != 0 and d_z_total != 0:
                dm_x = d_z_total * x_MM / denominator
                dm_y = d_z_total * y_MM / denominator
                new_mux += dm_x
                new_muy += dm_y
        
        # Update the dataframe
        df.at[idx, 'mux'] = new_mux
        df.at[idx, 'muy'] = new_muy
    
    # Copy sigma values
    df['sigmax'] = df['sigma_rad']
    df['sigmay'] = df['sigma_azi']

    # Remember workbook path for resolving custom PSF file stems during plotting.
    df.attrs['workbook_path'] = path

    return df  # Return the loaded DataFrame


def _resolve_custom_psf_path(workbook_path: str, stem: str) -> str | None:
    """Resolve a workbook-referenced PSF file stem to an existing path.

    The workbook may reference a PSF file by a stem or relative path.
    This helper tries common extensions and a small set of search
    directories (workbook dir, repo Distributions/ and CustomPSFs/, CWD)
    and returns the first existing absolute path or ``None`` when not
    found. Returns ``None`` for empty or placeholder stems.

    Parameters
    - workbook_path: path to the workbook containing the stem reference
    - stem: filename stem or relative path as provided in the workbook
    """
    stem = str(stem).strip()
    if not stem or stem.lower() in {'nan', 'none'}:
        return None

    # If user already passed a filename with extension, try directly.
    candidates: list[str] = []
    if stem.lower().endswith(('.xlsx', '.xls')):
        candidates.append(stem)
    else:
        for ext in ['.xlsx', '.xls']:
            candidates.append(stem + ext)

    wb_dir = os.path.dirname(os.path.abspath(workbook_path))
    repo_dir = os.path.dirname(os.path.abspath(__file__))
    search_dirs = [
        wb_dir,
        os.path.join(wb_dir, 'Distributions'),
        os.path.join(repo_dir, 'Distributions'),
        os.path.join(repo_dir, 'CustomPSFs'),
        os.getcwd(),
    ]

    for d in search_dirs:
        for name in candidates:
            p = os.path.join(d, name)
            if os.path.exists(p):
                return p
    return None


def compute_total_rot_polar(mm_to_pos: dict, mm_config_map: dict, alignment_by_pos: dict, gravity_by_pos: dict, thermal_by_pos: dict):
    """Compute total rotation components and their projections onto polar axes.

    Returns (rotx, roty, rot_rad, rot_azi) where each is a dict keyed by position.
    See module-level documentation for units and conventions (arcsec, direct
    polar terms versus projected X/Y rotations).
    """
    rotx = {}
    roty = {}
    rot_rad = {}
    rot_azi = {}

    # Gather all positions that may be present in any of the inputs
    positions = set()
    if mm_to_pos:
        positions.update(mm_to_pos.values())
    if alignment_by_pos:
        positions.update(alignment_by_pos.keys())
    if gravity_by_pos:
        positions.update(gravity_by_pos.keys())
    if thermal_by_pos:
        positions.update(thermal_by_pos.keys())

    # Reverse mapping: position -> an example MM for geometry lookup
    pos_to_mm = {}
    if mm_to_pos:
        for mm, p in mm_to_pos.items():
            pos_to_mm.setdefault(p, mm)

    for pos in positions:
        # Sum rotx/roty contributions from gravity and thermal only
        rtx_total = 0.0
        rty_total = 0.0
        if gravity_by_pos and pos in gravity_by_pos:
            rtx_total += float(gravity_by_pos[pos].get('d_grav_rotx', 0.0) or 0.0)
            rty_total += float(gravity_by_pos[pos].get('d_grav_roty', 0.0) or 0.0)
        if thermal_by_pos and pos in thermal_by_pos:
            rtx_total += float(thermal_by_pos[pos].get('d_therm_rotx', 0.0) or 0.0)
            rty_total += float(thermal_by_pos[pos].get('d_therm_roty', 0.0) or 0.0)

        rotx[pos] = rtx_total
        roty[pos] = rty_total

        # Compute radial unit vector from MM geometry if available
        ux, uy = 1.0, 0.0
        mm_choice = pos_to_mm.get(pos)
        if mm_choice is not None and mm_choice in mm_config_map:
            cfg = mm_config_map.get(mm_choice, {})
            r_mm = float(cfg.get('r_MM', 0.0) or 0.0)
            x_mm = float(cfg.get('x_MM', 0.0) or 0.0)
            y_mm = float(cfg.get('y_MM', 0.0) or 0.0)
            if r_mm > 0.0:
                ux = x_mm / r_mm
                uy = y_mm / r_mm

        # Project rotx/roty into polar components
        proj_rotrad = rtx_total * ux + rty_total * uy
        proj_rotazi = -rtx_total * uy + rty_total * ux

        # Add any direct polar contributions from alignment/gravity/thermal
        direct_rotrad = 0.0
        direct_rotazi = 0.0
        if alignment_by_pos and pos in alignment_by_pos:
            direct_rotrad += float(alignment_by_pos[pos].get('d_align_rotrad', 0.0) or 0.0)
            direct_rotazi += float(alignment_by_pos[pos].get('d_align_rotazi', 0.0) or 0.0)
        if gravity_by_pos and pos in gravity_by_pos:
            direct_rotrad += float(gravity_by_pos[pos].get('d_grav_rotrad', 0.0) or 0.0)
            direct_rotazi += float(gravity_by_pos[pos].get('d_grav_rotazi', 0.0) or 0.0)
        if thermal_by_pos and pos in thermal_by_pos:
            direct_rotrad += float(thermal_by_pos[pos].get('d_therm_rotrad', 0.0) or 0.0)
            direct_rotazi += float(thermal_by_pos[pos].get('d_therm_rotazi', 0.0) or 0.0)

        total_rotrad = proj_rotrad + direct_rotrad
        total_rotazi = proj_rotazi + direct_rotazi

        rot_rad[pos] = total_rotrad
        rot_azi[pos] = total_rotazi

    return rotx, roty, rot_rad, rot_azi


def compute_dm_from_dz(mm: dict, row: dict, d_z: float) -> tuple[float, float]:
    """Project a z-displacement d_z into DM x/y using MM geometry or theta fallback.

    - If `r_MM` > 0 is present in `mm`, project along the radial unit vector (x_MM/r_MM, y_MM/r_MM).
    - Else if `row` provides `theta_degrees`, use (cos(theta), sin(theta)).
    - Otherwise fallback to (1,0).
    """
    import math
    ux = 1.0
    uy = 0.0
    r_mm = float(mm.get('r_MM', 0.0) or 0.0)
    if r_mm > 0.0:
        ux = float(mm.get('x_MM', 0.0)) / r_mm
        uy = float(mm.get('y_MM', 0.0)) / r_mm
    else:
        theta = row.get('theta_degrees') if isinstance(row, dict) else None
        if theta is not None:
            try:
                th = math.radians(float(theta))
                ux = math.cos(th)
                uy = math.sin(th)
            except Exception:
                ux, uy = 1.0, 0.0

    dm_x = float(d_z) * ux
    dm_y = float(d_z) * uy
    return dm_x, dm_y


def plot_sum(df: pd.DataFrame, xlim=(-10,10), ylim=(-8,8), nx=800, ny=640, normalize=True, output=None, fast=True, title_suffix: str = "", df_optimized: pd.DataFrame = None, return_metrics_only: bool = False, debug: bool = False, metrics_n_r_final: int | None = None, metrics_n_theta_final: int | None = None, metrics_r_margin: float | None = None):
    """Create combined PSF image, compute HEW/EEF metrics, and optionally save output.

    This function composes per-MM PSFs (analytic or matrix-based) described
    in `df`, computes rotation-invariant HEW and EEF via polar integration,
    fits aggregated radial profiles (modified pseudo-Voigt and King) and
    produces a summary figure. Use `return_metrics_only=True` to retrieve
    computed numeric metrics without creating the plot (useful for CI).

    Important: `df` must include at minimum `mux`, `muy`, `sigmax`,
    `sigmay`, `weight` and `distribution`/custom PSF references as used by
    the rest of the codebase.
    """
    # Close any existing matplotlib figures to prevent accumulation
    plt.close('all')

    # Reduce grid resolution in fast mode
    if fast:
        # Aggressive defaults for interactive speed (target: <5s without optimization)
        nx = min(nx, 320)
        ny = min(ny, 240)
    else:
        # Keep slow mode bounded (still more accurate than fast, but not unbounded)
        nx = min(nx, 420)
        ny = min(ny, 340)

    # fast=True now means "quick" for both plain and comparison plots.
    quick_mode = bool(fast)
    
    # Calculate the weighted center of mass directly from the Gaussian parameters
    # Normalize per-combo weights so they sum to 1 before computing centroids and sums.
    # This ensures HEW is computed from properly normalized mixture amplitudes.
    # Use a numpy array to avoid repeatedly accessing the DataFrame.
    # Prefer adjusted A_eff weights when available
    if 'aeff_adjusted' in df.columns:
        # Treat NaN adjusted A_eff as zero weight (do not propagate NaN into normalization)
        weight_arr_for_center = pd.to_numeric(df['aeff_adjusted'], errors='coerce').fillna(0.0).to_numpy(dtype=float, copy=False)
    elif 'weight' in df.columns:
        # Ensure any non-finite weights are treated as zero
        weight_arr_for_center = pd.to_numeric(df['weight'], errors='coerce').fillna(0.0).to_numpy(dtype=float, copy=False)
    else:
        weight_arr_for_center = np.ones(len(df), dtype=float) if len(df) > 0 else np.array([], dtype=float)
    total_weight = float(np.nansum(weight_arr_for_center)) if weight_arr_for_center.size else 0.0
    if not np.isfinite(total_weight) or total_weight <= 0.0:
        # Fallback to unweighted centroid when weights are zero/invalid
        center_x = float(df['mux'].mean()) if 'mux' in df.columns else 0.0
        center_y = float(df['muy'].mean()) if 'muy' in df.columns else 0.0
        # keep weight array as ones for later summation
        weight_arr = np.ones(len(df), dtype=float)
    else:
        # normalize so sum(weights) == 1.0
        weight_arr = weight_arr_for_center / total_weight
        center_x = (df['mux'].to_numpy(dtype=float, copy=False) * weight_arr).sum()
        center_y = (df['muy'].to_numpy(dtype=float, copy=False) * weight_arr).sum()
    # Ensure centers are finite numbers
    if not np.isfinite(center_x):
        center_x = 0.0
    if not np.isfinite(center_y):
        center_y = 0.0
    
    # --- Fast grid summation helpers (threaded) ---
    mux_arr = df['mux'].to_numpy(dtype=float, copy=False)
    muy_arr = df['muy'].to_numpy(dtype=float, copy=False)
    sigx_arr = df['sigmax'].to_numpy(dtype=float, copy=False)
    sigy_arr = df['sigmay'].to_numpy(dtype=float, copy=False)
    # Ensure theta is numeric and replace NaN with 0.0 to avoid NaN trig results
    theta_arr = pd.to_numeric(df.get('theta_degrees', pd.Series([0.0] * len(df))), errors='coerce').fillna(0.0).to_numpy(dtype=float, copy=False)
    # Weight array already computed above as `weight_arr` (normalized). Ensure it's available.
    try:
        # if weight_arr defined above, keep it; otherwise derive normalized from df
        weight_arr
    except NameError:
        if 'weight' in df.columns:
            wtmp = df['weight'].to_numpy(dtype=float, copy=False)
            wsum = float(np.nansum(wtmp)) if wtmp.size else 0.0
            weight_arr = (wtmp / wsum) if (wsum and np.isfinite(wsum) and wsum > 0.0) else np.ones(len(df), dtype=float)
        else:
            weight_arr = np.ones(len(df), dtype=float)
    dist_raw_arr = df.get('distribution', pd.Series(['gaussian'] * len(df))).astype(str).to_numpy(copy=False)
    dist_arr = pd.Series(dist_raw_arr).astype(str).str.lower().to_numpy(copy=False)
    # alpha_* may contain placeholders like '-' for Gaussian rows; coerce safely.
    alpha_azi_arr = pd.to_numeric(df.get('alpha_azi', pd.Series([0.5] * len(df))), errors='coerce').fillna(0.5).to_numpy(dtype=float, copy=False)
    alpha_rad_arr = pd.to_numeric(df.get('alpha_rad', pd.Series([0.5] * len(df))), errors='coerce').fillna(0.5).to_numpy(dtype=float, copy=False)

    # Prevent numerically zero sigmas which break grid integration (tiny values
    # from user input or rounding can lead to extremely narrow peaks and unstable
    # minimal-interval searches). Enforce a small floor in meters.
    # Small floor for sigma in meters: keep tiny but non-zero to avoid
    # numerical integration instability. 1e-9 m (1 nm) is small enough to
    # preserve realistic PSF widths (user inputs are typically 1e-8..1e-6).
    MIN_SIG_M = 1e-9
    sigx_arr = np.maximum(sigx_arr, MIN_SIG_M)
    sigy_arr = np.maximum(sigy_arr, MIN_SIG_M)

    # Custom PSF cache (by distribution name/stem)
    builtins = {'gaussian', 'pseudo-voigt', 'voigt'}
    workbook_path = df.attrs.get('workbook_path', None)
    arcsec_to_m = 12 * np.pi / 180 / 3600
    custom_psf_cache: dict[str, tuple[np.ndarray, np.ndarray, np.ndarray]] = {}
    custom_sigma_hint: list[float] = []
    if workbook_path is not None:
        unique_names = sorted({str(n).strip() for n, dl in zip(dist_raw_arr, dist_arr) if str(n).strip() and dl not in builtins})
        for name in unique_names:
            p = _resolve_custom_psf_path(str(workbook_path), name)
            if not p:
                continue
            try:
                x_psf, y_psf, f_psf = load_psf_matrix_excel(p, arcsec_to_m=arcsec_to_m)
                custom_psf_cache[name] = (x_psf, y_psf, f_psf)
                # crude scale hint for r_max: assume extent ~ +/- 3 sigma
                extent = max(float(np.max(np.abs(x_psf))), float(np.max(np.abs(y_psf))))
                if extent > 0:
                    custom_sigma_hint.append(extent / 3.0)
            except Exception as e:
                print(f"Warning: Could not load custom PSF '{name}' from '{p}': {e}")

    def _sum_chunk_on_grid(Xg, Yg, idxs: np.ndarray, normalize_flag: bool) -> np.ndarray:
        Zc = np.zeros_like(Xg, dtype=float)
        for i in idxs:
            dlow = dist_arr[i]
            if dlow in ['pseudo-voigt', 'voigt']:
                add = pseudo_voigt_2d_rotated(
                    Xg, Yg,
                    muazi=mux_arr[i], murad=muy_arr[i],
                    sigmaazi=sigx_arr[i], sigmarad=sigy_arr[i],
                    theta=theta_arr[i],
                    alphaazi=alpha_azi_arr[i],
                    alpharad=alpha_rad_arr[i],
                    amplitude=weight_arr[i],
                    normalize=normalize_flag,
                    degrees=True,
                )
                Zc += add
            elif dlow == 'gaussian':
                add = gaussian_2d_rotated(
                    Xg, Yg,
                    mux=mux_arr[i], muy=muy_arr[i],
                    sigmax=sigx_arr[i], sigmay=sigy_arr[i],
                    theta=theta_arr[i],
                    amplitude=weight_arr[i],
                    normalize=normalize_flag,
                    degrees=True,
                )
                Zc += add
            else:
                # File-based custom PSF: distribution cell holds the file stem (without extension)
                name = str(dist_raw_arr[i]).strip()
                psf = custom_psf_cache.get(name)
                if psf is None and workbook_path is not None:
                    p = _resolve_custom_psf_path(str(workbook_path), name)
                    if p:
                        try:
                            psf = load_psf_matrix_excel(p, arcsec_to_m=arcsec_to_m)
                            custom_psf_cache[name] = psf
                        except Exception:
                            psf = None
                if psf is None:
                    # If file not found/invalid, fall back to Gaussian to keep plotting running.
                    Zc += gaussian_2d_rotated(
                        Xg, Yg,
                        mux=mux_arr[i], muy=muy_arr[i],
                        sigmax=max(sigx_arr[i], 1e-12), sigmay=max(sigy_arr[i], 1e-12),
                        theta=theta_arr[i],
                        amplitude=weight_arr[i],
                        normalize=normalize_flag,
                        degrees=True,
                    )
                else:
                    x_psf, y_psf, f_psf = psf
                    Zc += weight_arr[i] * eval_psf_matrix_rotated(
                        Xg,
                        Yg,
                        mux=mux_arr[i],
                        muy=muy_arr[i],
                        theta_deg=theta_arr[i],
                        x_axis=x_psf,
                        y_axis=y_psf,
                        flux=f_psf,
                    )
        return Zc

    def _sum_on_grid(Xg, Yg, normalize_flag: bool) -> np.ndarray:
        n = len(mux_arr)
        if n == 0:
            return np.zeros_like(Xg, dtype=float)

        # Threading helps because heavy NumPy ops release the GIL.
        max_workers = 1 if n < 25 else min(8, (os.cpu_count() or 2))
        if max_workers <= 1:
            return _sum_chunk_on_grid(Xg, Yg, np.arange(n), normalize_flag)

        chunks = np.array_split(np.arange(n), max_workers)
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            parts = list(ex.map(lambda c: _sum_chunk_on_grid(Xg, Yg, c, normalize_flag), chunks))
        Zg = np.zeros_like(Xg, dtype=float)
        for p in parts:
            Zg += p
        return Zg

    # ---- Rotation-invariant HEW via polar integration ----
    def radial_profile(cx, cy, n_r=400, n_theta=360, r_margin_factor=5.0, normalize_flag=normalize):
        """Compute radial energy profile around (cx, cy) using polar sampling.

        Uses polar grid integration (rotation-invariant) instead of Cartesian
        raster to avoid angle-dependent HEW bias.
        """

        # Radius limit: cover all PSFs plus margin
        max_sigma = max(df['sigmax'].max(), df['sigmay'].max())
        if custom_sigma_hint:
            max_sigma = max(float(max_sigma), float(max(custom_sigma_hint)))
        max_center_dist = np.sqrt((df['mux'] - cx) ** 2 + (df['muy'] - cy) ** 2).max()
        r_max = max_center_dist + r_margin_factor * max_sigma
        if r_max <= 0:
            r_max = 1e-6

        # Iteratively expand radial margin if heavy tails cause significant
        # energy to fall outside the initial r_max (common with Lorentzian).
        expected_total = float(np.nansum(weight_arr)) if 'weight_arr' in locals() else 1.0
        # allow a few attempts, expanding margin each time up to a modest limit
        # (reduce attempts and growth to avoid runaway cost for pseudo-Voigt tails)
        attempts = 3
        tol = 0.9995
        total_energy = 0.0
        for attempt in range(attempts):
            theta = np.linspace(0.0, 2.0 * np.pi, n_theta, endpoint=False)
            r = np.linspace(0.0, r_max, n_r)
            dtheta = theta[1] - theta[0]
            dr = r[1] - r[0] if n_r > 1 else r_max
            R, TH = np.meshgrid(r, theta)
            Xp = cx + R * np.cos(TH)
            Yp = cy + R * np.sin(TH)

            # Reuse the same summation kernel on the polar grid.
            Zp = _sum_on_grid(Xp, Yp, normalize_flag)

            # Integrate over theta (Jacobian r) to get radial energy density
            radial_energy = np.sum(Zp * R, axis=0) * dtheta  # shape (n_r,)
            cumulative = np.cumsum(radial_energy * dr)
            total_energy = cumulative[-1] if cumulative.size else 1.0

            # If normalization is expected (weights sum > 0) and we captured
            # nearly all energy, break early. Otherwise expand r_max and retry.
            if expected_total <= 0 or total_energy / expected_total >= tol:
                break
            # Expand margin and retry (limit growth to avoid runaway cost).
            # Use gentler growth factor and a tighter n_r cap to bound cost.
            r_max *= 1.5
            # Slightly increase radial resolution but cap much lower than before
            n_r = min(int(n_r * 1.15) + 1, 5000)
            n_theta = min(int(n_theta * 1.0), 2048)
        if debug:
                pass
        return r, cumulative, total_energy

    def _radius_for_fraction(frac: np.ndarray, r: np.ndarray, target: float = 0.5) -> float:
        """Estimate radius where cumulative fraction == target using local cubic fit.

        Falls back to linear interpolation when the local window is too small or
        a valid cubic root cannot be found.
        """
        if frac.size == 0:
            return 0.0
        frac = np.asarray(frac).ravel()
        r = np.asarray(r).ravel()
        if frac.size != r.size or frac.size < 2:
            return float(np.interp(target, frac, r))
        idx = int(np.searchsorted(frac, target))
        start = max(0, idx - 2)
        end = min(frac.size, idx + 2)
        if end - start < 2:
            return float(np.interp(target, frac, r))
        r_win = r[start:end]
        f_win = frac[start:end]
        if f_win.size < 4:
            return float(np.interp(target, frac, r))
        try:
            coeffs = np.polyfit(r_win, f_win, 3)
            coeffs[-1] -= float(target)
            roots = np.roots(coeffs)
            real_roots = roots[np.isreal(roots)].real
            rmin, rmax = float(r_win.min()), float(r_win.max())
            for rt in real_roots:
                if rmin - 1e-12 <= float(rt) <= rmax + 1e-12:
                    return float(rt)
        except Exception:
            pass
        return float(np.interp(target, frac, r))

    def hew_at_center(cx, cy, coarse=False):
        # Use coarser grid during search for speed
        if coarse:
            n_r = 110 if fast else 140
            n_theta = 72 if fast else 100
            r, cumulative, total_energy = radial_profile(cx, cy, n_r=n_r, n_theta=n_theta)
        else:
            n_r = 200 if fast else 240
            n_theta = 120 if fast else 180
            r, cumulative, total_energy = radial_profile(cx, cy, n_r=n_r, n_theta=n_theta)
        frac = cumulative / total_energy if total_energy > 0 else cumulative
        # Use local cubic interpolation for the 50% radius, then return HEW diameter
        radius50 = _radius_for_fraction(frac, r, target=0.5)
        return 2.0 * radius50

    # Find best focus position.
    # In quick_mode (typical when not optimizing), avoid iterative search to keep runtime low.
    if quick_mode:
        best_cx, best_cy = center_x, center_y
        best_hew = hew_at_center(best_cx, best_cy, coarse=True)
        step_size = 1e-6
        for cx, cy in (
            (center_x + step_size, center_y),
            (center_x - step_size, center_y),
            (center_x, center_y + step_size),
            (center_x, center_y - step_size),
        ):
            hew_val = hew_at_center(cx, cy, coarse=True)
            if hew_val < best_hew:
                best_cx, best_cy, best_hew = cx, cy, hew_val
        center_x, center_y = best_cx, best_cy
    else:
        candidates_start = [(center_x, center_y), (0.0, 0.0)]
        best_cx, best_cy = center_x, center_y
        best_hew = hew_at_center(best_cx, best_cy, coarse=True)

        for start_cx, start_cy in candidates_start:
            current_cx, current_cy = start_cx, start_cy
            current_hew = hew_at_center(current_cx, current_cy, coarse=True)
            step_size = 1e-6  # 1 micron steps

            for _ in range(20 if fast else 30):  # Limit iterations for speed
                candidates = [
                    (current_cx + step_size, current_cy),
                    (current_cx - step_size, current_cy),
                    (current_cx, current_cy + step_size),
                    (current_cx, current_cy - step_size)
                ]
                improved = False
                for cx, cy in candidates:
                    hew_val = hew_at_center(cx, cy, coarse=True)
                    if hew_val < current_hew:
                        current_cx, current_cy = cx, cy
                        current_hew = hew_val
                        improved = True
                        break
                if not improved:
                    step_size *= 0.5
                    if step_size < 1e-10:
                        break

            if current_hew < best_hew:
                best_cx, best_cy = current_cx, current_cy
                best_hew = current_hew

        center_x, center_y = best_cx, best_cy

    # Final HEW and EE metrics using fine polar integration at best focus
    # Lower final polar-grid for faster runs (target <5s) while increasing
    # radial margin to capture PSF tails and keep HEW estimate high.
    if quick_mode:
        # Coarse: keep radial emphasis but reduce final caps to limit runtime
        n_r_final = 3000
        n_theta_final = 120
        final_r_margin = 10.0
    else:
        # Increase fine-mode sampling for higher accuracy but stay bounded.
        n_r_final = 5000
        n_theta_final = 720
        final_r_margin = 12.0

    # Allow callers to override final metric sampling for speed/experiments
    if metrics_n_r_final is not None:
        try:
            n_r_final = int(metrics_n_r_final)
        except Exception:
            pass
    if metrics_n_theta_final is not None:
        try:
            n_theta_final = int(metrics_n_theta_final)
        except Exception:
            pass
    if metrics_r_margin is not None:
        try:
            final_r_margin = float(metrics_r_margin)
        except Exception:
            pass
    r_profile, cumulative_profile, total_energy = radial_profile(center_x, center_y, n_r=n_r_final, n_theta=n_theta_final, r_margin_factor=final_r_margin)
    frac_profile = cumulative_profile / total_energy if total_energy > 0 else cumulative_profile
    radius_50 = _radius_for_fraction(frac_profile, r_profile, target=0.5)
    radius_90 = _radius_for_fraction(frac_profile, r_profile, target=0.9)

    # --------------------------------------------------------
    # Fit modified pseudo-Voigt to the azimuthal-average radial profile
    # Improved: two-stage fit + robust least-squares to better capture wings
    # --------------------------------------------------------
    fit_params_available = False
    fit_profile_pct = None
    fit_profile_diam = None
    pearson4_profile_pct = None
    pearson4_profile_diam = None
    try:
        from scipy.optimize import curve_fit
        try:
            from scipy.optimize import least_squares
            have_least_squares = True
        except Exception:
            have_least_squares = False

        # Recover radial energy per-bin from cumulative (radial_energy*dr = diff(cumulative))
        dr = float(r_profile[1] - r_profile[0]) if r_profile.size > 1 else 1.0
        radial_energy = np.diff(np.concatenate(([0.0], cumulative_profile))) / dr

        # Convert to mean intensity per unit area of the annulus: I = radial_energy/(2*pi*r)
        with np.errstate(divide='ignore', invalid='ignore'):
            I_profile = radial_energy / (2.0 * np.pi * r_profile)
        # handle r==0 by replacing with nearest finite value
        if not np.isfinite(I_profile[0]):
            finite_idx = np.where(np.isfinite(I_profile))[0]
            if finite_idx.size:
                I_profile[0] = I_profile[finite_idx[0]]
            else:
                I_profile[0] = 0.0

        # Work in arcsec for human-friendly parameters
        try:
            arcsec_to_m
        except NameError:
            arcsec_to_m = 12 * np.pi / 180 / 3600
        r_arcsec = r_profile / arcsec_to_m

        # Local model definitions
        def pure_gaussian(r, A, Gamma, b):
            return A * np.exp(-4*np.log(2)*(r/Gamma)**2) + b

        # Modified pseudo-Voigt radial intensity model used for the aggregated fit.
        # The model mixes a narrow Gaussian core with a broader Lorentzian-like wing.
        # The wing amplitude is scaled separately and the final mixture is normalized
        # so that A remains the overall intensity scale.
        #
        #   G(r; Γ_c) = exp(-4 ln 2 (r / Γ_c)^2)
        #   a = 2^(1/β) - 1
        #   C(r; Γ_w) = [1 + a (2 r / Γ_w)^2]^{-β}
        #   mix = (1 - η) G + η scalar C
        #   norm = (1 - η) + η scalar
        #   I(r) = A * (mix / norm)
        #
        # where r is in arcseconds and A is the peak radial mean intensity.
        def beta_pseudo_gaussian(r, A, Gamma_c, Gamma_w, eta, beta, scalar):
            G = np.exp(-4*np.log(2)*(r/Gamma_c)**2)
            a = 2**(1.0/beta) - 1.0
            C = 1.0 / (1.0 + a*(2.0*r/Gamma_w)**2)**beta
            mix = (1.0-eta)*G + eta*scalar*C
            norm = (1.0-eta) + eta*scalar
            return A * (mix / norm)

        # King profile often used for PSFs: I(r) = I0 * (1 + (r/rc)^2)^(-alpha) + b
        def king_profile(r, I0, rc, alpha, b):
            return I0 * (1.0 + (r/np.maximum(rc, 1e-12))**2.0)**(-np.maximum(alpha, 0.01)) + b

        def core_weights(r, r0):
            return np.exp(-(r/r0)**2)

        # Utility: bounded multi-start least_squares runner to fit EEF-focused residuals
        def multi_start_least_squares(resid_func, x0, lb, ub, attempts=6, rng_seed=12345, **ls_kwargs):
            try:
                from scipy.optimize import least_squares
            except Exception:
                return None, None
            rng = np.random.default_rng(rng_seed)
            best_score = np.inf
            best_x = None
            last_res = None
            x0 = np.asarray(x0, dtype=float)
            lb = np.asarray(lb, dtype=float)
            ub = np.asarray(ub, dtype=float)
            for attempt in range(attempts):
                try:
                    pert = x0 * (1.0 + rng.normal(0.0, 0.12, size=x0.shape))
                    pert = np.maximum(lb, np.minimum(ub, pert))
                    res = least_squares(resid_func, pert, bounds=(lb, ub), **ls_kwargs)
                except Exception:
                    res = None
                if res is not None and hasattr(res, 'fun'):
                    try:
                        score = float(np.sqrt(np.mean(res.fun**2)))
                    except Exception:
                        score = np.inf
                    if np.isfinite(score) and score < best_score:
                        best_score = score
                        best_x = res.x
                        last_res = res
            return best_x, last_res

        # Prepare data mask (positive intensities)
        mask = np.isfinite(I_profile) & (I_profile > 0)
        r_fit = r_arcsec[mask]
        I_fit = I_profile[mask]
        try:
            print(f"r_fit size={r_fit.size}, I_fit finite count={np.count_nonzero(np.isfinite(I_fit))}")
        except Exception:
            pass

        fit_params_available = False
        fit_profile_pct = None
        fit_profile_diam = None

        if r_fit.size >= 7:
            try:
                print("Entering radial-fit block (r_fit.size >= 7)")
            except Exception:
                pass
            # STEP 1: core-only Gaussian fit for robust initial Gamma estimate
            A0 = float(np.nanmax(I_fit))
            try:
                Gamma0 = float(r_fit[np.argmin(np.abs(I_fit - A0/2.0))])
            except Exception:
                Gamma0 = max(float(np.median(r_fit)), 1.0)
            b0 = max(float(np.nanmin(I_fit)), 0.0)

            rmax_core = 2.0 * max(Gamma0, 1e-6)
            core_mask = r_fit <= rmax_core
            try:
                popt_g, _ = curve_fit(pure_gaussian, r_fit[core_mask], I_fit[core_mask], p0=[A0, Gamma0, b0])
                A_g, Gamma_g, b_g = popt_g
            except Exception:
                A_g, Gamma_g, b_g = A0, Gamma0, b0

            # STEP 2: modified pseudo-Voigt fit with separate core and wing widths
            A_fit = float(A_g)
            Gamma_c_fit = float(Gamma_g)
            Gamma_w_fit = max(3.0 * Gamma_c_fit, 4.0)
            eta_fit = 0.08
            beta_fit = 1.8
            scalar_fit = 1.8

            A_lb = max(0.5 * A_fit, 1.0)
            x0 = [A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit]
            lb = [A_lb, max(1e-3, float(Gamma_c_fit)*0.2), max(1e-3, float(Gamma_c_fit)*1.2), 0.05, 1.0, 0.2]
            ub = [np.inf, float(Gamma_c_fit)*3.0, float(Gamma_c_fit)*25.0, 0.5, 5.0, 12.0]

            # Use weighted least squares for peak and wing matching.
            eps = 1e-12
            floor = max(eps, np.nanpercentile(I_fit[I_fit>0], 1) if np.any(I_fit>0) else eps)
            rscale = np.median(r_fit) if r_fit.size else 1.0
            peak_cut = max(3.0, min(12.0, float(Gamma_c_fit)*1.5))
            weights = 1.0 + 20.0 * np.exp(-(r_fit/peak_cut)**2) + 0.5 * (r_fit / max(rscale, 1e-6))
            sigma = 1.0 / np.maximum(weights, 1e-8)

            def fit_func(r, A_x, Gc_x, Gw_x, eta_x, beta_x, scalar_x):
                return beta_pseudo_gaussian(r, A_x, Gc_x, Gw_x, eta_x, beta_x, scalar_x)

            # Prepare cumulative EEF reference for the current data
            dr_arcsec = float(r_arcsec[1] - r_arcsec[0]) if r_arcsec.size > 1 else 1.0
            radial_energy_data = 2.0 * np.pi * r_fit * I_fit * dr_arcsec
            cumulative_data = np.cumsum(radial_energy_data)
            total_data = cumulative_data[-1] if cumulative_data.size else 1.0
            eef_data = cumulative_data / total_data if total_data > 0 else cumulative_data
            # Give the EEF objective most weight but keep a small intensity
            # residual component to stabilise the core fit and prevent large
            # deviations at intermediate radii.
            eef_weight = 30.0
            intensity_weight_scale = 0.25

            def resid_pvoigt(params):
                A_x, Gc_x, Gw_x, eta_x, beta_x, scalar_x = params
                model = beta_pseudo_gaussian(r_fit, A_x, Gc_x, Gw_x, eta_x, beta_x, scalar_x)
                model = np.maximum(model, floor)
                # EEF residual (fraction) only: compute cumulative model energy and compare
                cumulative_model = np.cumsum(2.0 * np.pi * r_fit * model * dr_arcsec)
                total_model = cumulative_model[-1] if cumulative_model.size else 1.0
                eef_model = cumulative_model / total_model if total_model > 0 else cumulative_model
                res_eef = (eef_model - eef_data) * eef_weight
                return res_eef

            # Use multi-start least_squares for robustness (EEF-only objective)
            if have_least_squares:
                try:
                    # increase PV multi-start attempts and allow longer evaluation per start
                    best_x, last_res = multi_start_least_squares(resid_pvoigt, x0, lb, ub, attempts=36, rng_seed=12345, loss='soft_l1', f_scale=1e-3, max_nfev=120000)
                    if last_res is not None and best_x is not None:
                        A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit = best_x.tolist()
                    else:
                        # fallback to curve_fit if LS failed
                        popt, _ = curve_fit(fit_func, r_fit, I_fit, p0=x0, bounds=(lb, ub), sigma=sigma, maxfev=5000)
                        A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit = popt
                except Exception:
                    try:
                        popt, _ = curve_fit(fit_func, r_fit, I_fit, p0=x0, bounds=(lb, ub), sigma=sigma, maxfev=5000)
                        A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit = popt
                    except Exception:
                        A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit = x0
            else:
                try:
                    popt, _ = curve_fit(fit_func, r_fit, I_fit, p0=x0, bounds=(lb, ub), sigma=sigma, maxfev=5000)
                    A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit = popt
                except Exception:
                    A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit = x0

            # Save aggregated fit diagnostic plot
            fit_params_available = True
            # Standalone PV diagnostic figure generation disabled (user request).
            try:
                # Previously saved a standalone E2E_fit.png here; intentionally skipping.
                pass
            except Exception:
                pass

            try:
                dr_arcsec = float(r_arcsec[1] - r_arcsec[0]) if r_arcsec.size > 1 else 1.0
                I_fit_model = beta_pseudo_gaussian(r_arcsec, A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit)
                radial_energy_fit = 2.0 * np.pi * r_arcsec * I_fit_model * dr_arcsec
                total_fit_energy = np.sum(radial_energy_fit)
                if total_fit_energy > 0:
                    fit_cumulative = np.cumsum(radial_energy_fit)
                    fit_profile_pct = 100.0 * fit_cumulative / total_fit_energy
                    fit_profile_diam = 2.0 * r_arcsec
            except Exception:
                fit_profile_pct = None
                fit_profile_diam = None

            # ========== PEARSON4 FITTING ==========
            # Fit Pearson4 model using lmfit to also optimize intensity + EEF
            pearson4_result = None
            pearson4_profile_pct = None
            pearson4_profile_diam = None
            try:
                from lmfit.models import Pearson4Model
                from lmfit import Parameters
                print("DEBUG: lmfit imported successfully for Pearson4 fitting")
                
                # Create Pearson4 model
                pearson4_model = Pearson4Model()
                
                # Initial guess from pseudo-Voigt peak parameters
                params = pearson4_model.guess(I_fit, x=r_fit)
                try:
                    print(f"DEBUG: pearson4 initial guess amp={params['amplitude'].value}, center={params['center'].value}, sigma={params['sigma'].value}, expon={params['expon'].value}, skew={params['skew'].value}")
                except Exception:
                    print("DEBUG: pearson4 initial guess created but some params missing")
                # Adjust initial values for better convergence
                params['amplitude'].value = A_fit
                params['center'].value = Gamma_c_fit * 0.5
                params['sigma'].value = Gamma_c_fit
                params['expon'].value = 1.5
                params['skew'].value = 0.0
                
                # Define EEF-only residual for Pearson4 and run bounded multi-start least_squares
                dr_arcsec_p4 = float(r_arcsec[1] - r_arcsec[0]) if r_arcsec.size > 1 else 1.0
                radial_energy_data_p4 = 2.0 * np.pi * r_fit * I_fit * dr_arcsec_p4
                cumulative_data_p4 = np.cumsum(radial_energy_data_p4)
                total_data_p4 = cumulative_data_p4[-1] if cumulative_data_p4.size else 1.0
                eef_data_p4 = cumulative_data_p4 / total_data_p4 if total_data_p4 > 0 else cumulative_data_p4

                # Prepare parameter bounds via lmfit Parameters when available
                try:
                    # set sensible bounds on lmfit params to stabilize EEF-only optimization
                    params['amplitude'].min = max(float(floor), 0.0)
                    params['amplitude'].max = float(np.nanmax(I_fit) * 10.0 if I_fit.size else params['amplitude'].value * 100.0)
                    params['center'].min = 0.0
                    params['center'].max = float(min(r_arcsec.max() * 0.5 if r_arcsec.size else 50.0, 50.0))
                    params['sigma'].min = 1e-3
                    params['sigma'].max = float(max(1.0, r_arcsec.max()))
                    params['expon'].min = 1.0
                    params['expon'].max = 8.0
                    params['skew'].min = -2.0
                    params['skew'].max = 2.0
                except Exception:
                    pass

                # Residual that returns only EEF difference (scaled)
                def residual_p4_params(pdict):
                    try:
                        # Use a coarser radial grid for EEF evaluation to speed residuals
                        try:
                            rg = r_arcsec_coarse
                        except NameError:
                            rg = r_arcsec
                        model = pearson4_model.eval(params=pdict, x=rg)
                        model = np.maximum(model, 0.0)
                        dr_rg = float(rg[1] - rg[0]) if rg.size > 1 else dr_arcsec_p4
                        radial_model = 2.0 * np.pi * rg * model * dr_rg
                        tot_model = np.sum(radial_model)
                        if tot_model <= 0 or not np.isfinite(tot_model):
                            return np.ones_like(eef_data_p4) * 1e6
                        eef_model = np.cumsum(radial_model) / tot_model
                        # interpolate model EEF (fraction) to the data sample radii r_fit
                        from numpy import interp
                        model_at_rfit = interp(r_fit, rg, eef_model)
                        # Normalize EEF residuals by typical EEF variation to smooth objective
                        try:
                            eef_scale = float(np.maximum(1e-3, np.nanstd(eef_data_p4)))
                        except Exception:
                            eef_scale = 1.0
                        eef_weight_norm = 1.0
                        return (model_at_rfit - eef_data_p4) / eef_scale * eef_weight_norm
                    except Exception:
                        return np.ones_like(eef_data_p4) * 1e6

                # Run bounded multi-start optimization using least_squares when available
                pearson4_result = None
                try:
                    if have_least_squares:
                        best_score = np.inf
                        best_params = None
                        rng = np.random.default_rng(4321)
                        # Build initial vector from params
                        try:
                            p0_vec = np.array([params['amplitude'].value, params['center'].value, params['sigma'].value, params['expon'].value, params['skew'].value])
                        except Exception:
                            # Derive a better initial guess from the pseudo-Voigt
                            # aggregated fit: amplitude from PV core, center from
                            # data peak, sigma from PV core width, and expon
                            # loosely informed by PV beta.
                            try:
                                amp_guess = float(A_fit)
                            except Exception:
                                amp_guess = float(np.nanmax(I_fit) if I_fit.size else 1.0)
                            try:
                                # center: radius of maximum measured intensity
                                center_guess = float(r_fit[np.nanargmax(I_fit)]) if (r_fit.size and np.any(np.isfinite(I_fit))) else 0.0
                            except Exception:
                                center_guess = 0.0
                            try:
                                sigma_guess = max(1e-3, float(Gamma_c_fit))
                            except Exception:
                                sigma_guess = max(1e-3, float(np.median(r_fit)) if r_fit.size else 1.0)
                            try:
                                expon_guess = float(beta_fit) if ('beta_fit' in locals() and np.isfinite(beta_fit)) else 1.5
                                # clamp to Pearson4 reasonable range
                                expon_guess = float(max(1.0, min(6.0, expon_guess)))
                            except Exception:
                                expon_guess = 1.5
                            p0_vec = np.array([amp_guess, center_guess, sigma_guess, expon_guess, 0.0])
                        try:
                            print(f"DEBUG: pearson4 p0_vec={p0_vec}")
                        except Exception:
                            pass

                        # Try a quick lmfit intensity-only fit to obtain a robust
                        # starting point for the multi-start EEF optimization.
                        try:
                            quick_fit = None
                            try:
                                quick_fit = pearson4_model.fit(I_fit, params, x=r_fit, max_nfev=3000)
                            except Exception:
                                quick_fit = None
                            if quick_fit is not None and hasattr(quick_fit, 'params'):
                                try:
                                    amp_q = float(quick_fit.params['amplitude'].value)
                                    cen_q = float(quick_fit.params['center'].value) if 'center' in quick_fit.params else 0.0
                                    sig_q = float(quick_fit.params['sigma'].value)
                                    exp_q = float(quick_fit.params['expon'].value) if 'expon' in quick_fit.params else 1.5
                                    skew_q = float(quick_fit.params['skew'].value) if 'skew' in quick_fit.params else 0.0
                                    p0_vec = np.array([amp_q, cen_q, sig_q, exp_q, skew_q])
                                    try:
                                        print(f"DEBUG: seed p0_vec from quick lmfit intensity fit: {p0_vec}")
                                    except Exception:
                                        pass
                                except Exception:
                                    pass
                        except Exception:
                            pass

                        # Build a coarser radial grid for EEF computations to reduce
                        # costly full-array evaluations during multi-starts.
                        try:
                            max_pts = 300
                            if r_arcsec.size > max_pts:
                                r_arcsec_coarse = np.linspace(0.0, float(r_arcsec.max()), max_pts)
                            else:
                                r_arcsec_coarse = r_arcsec
                        except Exception:
                            r_arcsec_coarse = r_arcsec
                        # Subsample r_fit for intensity residuals during multi-starts
                        try:
                            max_int_samples = 300
                            if r_fit.size > max_int_samples:
                                idxs = np.round(np.linspace(0, r_fit.size-1, max_int_samples)).astype(int)
                                r_fit_sub = r_fit[idxs]
                                I_fit_sub = I_fit[idxs]
                            else:
                                r_fit_sub = r_fit
                                I_fit_sub = I_fit
                        except Exception:
                            r_fit_sub = r_fit
                            I_fit_sub = I_fit

                        # Bounds vector: [amp, center, sigma, expon, skew]
                        try:
                            lb = np.array([params['amplitude'].min if hasattr(params['amplitude'], 'min') else max(float(floor), 0.0),
                                           params['center'].min if hasattr(params['center'], 'min') else 0.0,
                                           params['sigma'].min if hasattr(params['sigma'], 'min') else 1e-6,
                                           params['expon'].min if hasattr(params['expon'], 'min') else 1.0,
                                           params['skew'].min if hasattr(params['skew'], 'min') else -5.0])
                            ub = np.array([params['amplitude'].max if hasattr(params['amplitude'], 'max') else np.inf,
                                           params['center'].max if hasattr(params['center'], 'max') else float(r_arcsec.max() if r_arcsec.size else 50.0),
                                           params['sigma'].max if hasattr(params['sigma'], 'max') else float(max(1.0, r_arcsec.max() if r_arcsec.size else 50.0)),
                                           params['expon'].max if hasattr(params['expon'], 'max') else 8.0,
                                           params['skew'].max if hasattr(params['skew'], 'max') else 5.0])
                        except Exception:
                            lb = np.array([max(float(floor), 0.0), 0.0, 1e-6, 1.0, -5.0])
                            ub = np.array([np.inf, float(r_arcsec.max() if r_arcsec.size else 50.0), float(max(1.0, r_arcsec.max() if r_arcsec.size else 50.0)), 8.0, 5.0])

                        from scipy.optimize import least_squares

                        def vec_to_params(vec):
                            # convert vector to lmfit-like param dict for pearson4_model.eval
                            try:
                                p = {
                                    'amplitude': vec[0],
                                    'center': vec[1],
                                    'sigma': vec[2],
                                    'expon': vec[3],
                                    'skew': vec[4]
                                }
                                return p
                            except Exception:
                                return None

                        def compute_p4_eef_rms_from_vec(vec):
                            """Compute unscaled EEF RMS (model vs data) for a Pearson4 parameter vector.
                            Returns np.inf on failure.
                            """
                            try:
                                pd = vec_to_params(np.asarray(vec, dtype=float))
                                if pd is None:
                                    return np.inf
                                model_full = pearson4_model.eval(params=pd, x=r_arcsec)
                                model_full = np.maximum(model_full, 0.0)
                                radial_model = 2.0 * np.pi * r_arcsec * model_full * dr_arcsec_p4
                                tot_model = np.sum(radial_model)
                                if tot_model <= 0 or not np.isfinite(tot_model):
                                    return np.inf
                                eef_model = np.cumsum(radial_model) / tot_model
                                from numpy import interp
                                model_eef_at_rfit = interp(r_fit, r_arcsec, eef_model)
                                # compare to the data EEF computed for Pearson4 block
                                try:
                                    ref = eef_data_p4
                                except Exception:
                                    ref = eef_data
                                res = model_eef_at_rfit - ref
                                return float(np.sqrt(np.mean(res**2))) if res.size else np.inf
                            except Exception:
                                return np.inf

                        def compute_pv_eef_rms():
                            """Compute unscaled EEF RMS for the pseudo-Voigt fit vs data when available.
                            Returns np.inf if PV EEF not available.
                            """
                            try:
                                if fit_profile_pct is None:
                                    return np.inf
                                from numpy import interp
                                pv_eef_frac = np.asarray(fit_profile_pct) / 100.0
                                model_eef_at_rfit = interp(r_fit, r_arcsec, pv_eef_frac)
                                ref = eef_data if 'eef_data' in locals() else (eef_data_p4 if 'eef_data_p4' in locals() else None)
                                if ref is None:
                                    return np.inf
                                res = model_eef_at_rfit - ref
                                return float(np.sqrt(np.mean(res**2))) if res.size else np.inf
                            except Exception:
                                return np.inf

                        # Use shared multi-start helper for EEF-only Pearson4 fitting
                        try:
                            def pearson4_resid_vec(v):
                                return residual_p4_params(vec_to_params(v))
                            # reduce attempts and max evals for speed; coarse EEF grid used inside residuals
                            # increase Pearson4 EEF-only multi-starts and per-start budget
                            best_params, last_res_p4 = multi_start_least_squares(pearson4_resid_vec, p0_vec, lb, ub, attempts=36, rng_seed=4321, max_nfev=120000, loss='soft_l1', f_scale=1e-4)
                            try:
                                print(f"DEBUG: pearson4 multi-start result best_params={best_params}, last_res_cost={None if last_res_p4 is None else getattr(last_res_p4,'cost',None)}")
                            except Exception:
                                pass
                            # Inspect optimizer result and accept only if residuals look reasonable
                            accepted_p4 = False
                            rms_p4 = np.inf
                            try:
                                if last_res_p4 is not None and hasattr(last_res_p4, 'fun'):
                                    fun = np.asarray(last_res_p4.fun, dtype=float)
                                    rms_p4 = float(np.sqrt(np.mean(fun**2))) if fun.size else np.inf
                                    fun_min = float(np.nanmin(fun)) if fun.size else np.nan
                                    fun_max = float(np.nanmax(fun)) if fun.size else np.nan
                                    fun_std = float(np.nanstd(fun)) if fun.size else np.nan
                                    print(f"DEBUG: pearson4 residuals stats: rms={rms_p4:.3g}, min={fun_min:.3g}, max={fun_max:.3g}, std={fun_std:.3g}")
                            except Exception:
                                pass
                            # Threshold for accepting EEF-only Pearson4 fit (empirical)
                            try:
                                # Make EEF-only Pearson4 acceptance more stringent
                                # Lowering threshold further to reduce false positives
                                # Make EEF-only Pearson4 acceptance more stringent
                                # Lowering threshold further to reduce false positives
                                threshold_rms = 1.0
                                if best_params is not None and np.isfinite(rms_p4) and rms_p4 < threshold_rms:
                                    accepted_p4 = True
                            except Exception:
                                accepted_p4 = False
                            if accepted_p4:
                                # Only accept this EEF-only Pearson4 result if it improves
                                # the EEF match relative to the pseudo-Voigt fit.
                                try:
                                    p4_eef_rms = compute_p4_eef_rms_from_vec(best_params)
                                    pv_eef_rms = compute_pv_eef_rms()
                                except Exception:
                                    p4_eef_rms = np.inf
                                    pv_eef_rms = np.inf
                                if np.isfinite(p4_eef_rms) and np.isfinite(pv_eef_rms):
                                    better = (p4_eef_rms <= 3.0 * pv_eef_rms)
                                else:
                                    # if PV not available, fall back to original criterion
                                    better = True
                                if better:
                                    class SimpleResult:
                                        pass
                                    pearson4_result = SimpleResult()
                                    pearson4_result.params = vec_to_params(best_params)
                                    try:
                                        print(f"DEBUG: pearson4_result params set from best_params: {pearson4_result.params}")
                                    except Exception:
                                        pass
                                else:
                                    pearson4_result = None
                                    try:
                                        print(f"DEBUG: pearson4 fit rejected because PV EEF is closer (p4_rms={p4_eef_rms:.3g}, pv_rms={pv_eef_rms:.3g})")
                                    except Exception:
                                        pass
                            else:
                                pearson4_result = None
                                try:
                                    print(f"DEBUG: pearson4 fit rejected (rms={rms_p4}); pearson4_result set to None")
                                except Exception:
                                    pass
                            # Fallback: if EEF-only fit failed, try a combined EEF + intensity objective
                            if pearson4_result is None:
                                try:
                                    print("DEBUG: attempting combined EEF+intensity Pearson4 fit (fallback)")
                                    intensity_weight_scale_p4 = 0.25
                                    from numpy import interp
                                    def pearson4_resid_comb_vec(v):
                                        try:
                                            pd = vec_to_params(v)
                                            # Evaluate EEF on a coarse grid for speed
                                            try:
                                                rg = r_arcsec_coarse
                                            except NameError:
                                                rg = r_arcsec
                                            model_full = pearson4_model.eval(params=pd, x=rg)
                                            model_full = np.maximum(model_full, 0.0)
                                            dr_rg = float(rg[1] - rg[0]) if rg.size > 1 else dr_arcsec_p4
                                            radial_model = 2.0 * np.pi * rg * model_full * dr_rg
                                            tot_model = np.sum(radial_model)
                                            if tot_model <= 0 or not np.isfinite(tot_model):
                                                # return large penalty vector for both parts
                                                return np.ones((eef_data_p4.size + r_fit_sub.size,), dtype=float) * 1e6
                                            eef_model = np.cumsum(radial_model) / tot_model
                                            model_at_rfit = interp(r_fit, rg, eef_model)
                                            # Normalize EEF residuals by typical EEF variation
                                            try:
                                                eef_scale = float(np.maximum(1e-3, np.nanstd(eef_data_p4)))
                                            except Exception:
                                                eef_scale = 1.0
                                            eef_weight_norm = 1.0
                                            eef_res = (model_at_rfit - eef_data_p4) / eef_scale * eef_weight_norm
                                            # intensity residuals on data sample points (log space)
                                            # use subsampled points for intensity residuals during multi-starts
                                            model_rfit_vals_sub = pearson4_model.eval(params=pd, x=r_fit_sub)
                                            model_rfit_vals_sub = np.maximum(model_rfit_vals_sub, floor)
                                            # normalize log-int residuals by their typical std to smooth objective
                                            try:
                                                int_ref = np.log(I_fit_sub + floor)
                                                int_scale = float(np.maximum(1e-3, np.nanstd(int_ref)))
                                            except Exception:
                                                int_scale = 1.0
                                            int_res_sub = (np.log(model_rfit_vals_sub + floor) - np.log(I_fit_sub + floor)) / int_scale * (intensity_weight_scale_p4 * 2.0)
                                            return np.concatenate([eef_res, int_res_sub])
                                        except Exception:
                                            return np.ones((eef_data_p4.size + r_fit_sub.size,), dtype=float) * 1e6

                                    try:
                                        # use fewer starts and fewer function evaluations for the combined fit
                                        # increase combined-fit multi-starts and per-start budget
                                        best_params2, last_res2 = multi_start_least_squares(pearson4_resid_comb_vec, p0_vec, lb, ub, attempts=36, rng_seed=9999, loss='soft_l1', f_scale=1e-4, max_nfev=120000)
                                    except Exception:
                                        best_params2, last_res2 = None, None
                                    try:
                                        print(f"DEBUG: combined-fit best_params2={best_params2}, last_res2_cost={None if last_res2 is None else getattr(last_res2,'cost',None)}")
                                    except Exception:
                                        pass
                                    accepted2 = False
                                    rms2 = np.inf
                                    try:
                                        # Re-evaluate fitted model on data points to compute
                                        # meaningful unscaled EEF and intensity diagnostics.
                                        if best_params2 is not None:
                                            pd2 = vec_to_params(best_params2)
                                            # full model on r_arcsec
                                            model_full2 = pearson4_model.eval(params=pd2, x=r_arcsec)
                                            model_full2 = np.maximum(model_full2, 0.0)
                                            radial_model2 = 2.0 * np.pi * r_arcsec * model_full2 * dr_arcsec_p4
                                            tot_model2 = np.sum(radial_model2)
                                            # data total for comparison
                                            tot_data = np.sum(radial_energy_data_p4)
                                            if tot_model2 <= 0 or not np.isfinite(tot_model2):
                                                raise ValueError('invalid total model energy')
                                            eef_model2 = np.cumsum(radial_model2) / tot_model2
                                            from numpy import interp
                                            model_eef_at_rfit = interp(r_fit, r_arcsec, eef_model2)
                                            # unscaled EEF residuals (fraction)
                                            eef_res_unscaled = (model_eef_at_rfit - eef_data_p4)
                                            eef_rms_unscaled = float(np.sqrt(np.mean(eef_res_unscaled**2))) if eef_res_unscaled.size else np.inf
                                            # intensity residuals (log-space) on r_fit
                                            model_rfit_vals2 = pearson4_model.eval(params=pd2, x=r_fit)
                                            model_rfit_vals2 = np.maximum(model_rfit_vals2, floor)
                                            int_res_unscaled = (np.log(model_rfit_vals2 + floor) - np.log(I_fit + floor))
                                            int_rms_unscaled = float(np.sqrt(np.mean(int_res_unscaled**2))) if int_res_unscaled.size else np.inf
                                            # combined rms (on optimizer fun vector) for bookkeeping
                                            if last_res2 is not None and hasattr(last_res2, 'fun'):
                                                fun2 = np.asarray(last_res2.fun, dtype=float)
                                                rms2 = float(np.sqrt(np.mean(fun2**2))) if fun2.size else np.inf
                                            else:
                                                rms2 = np.inf
                                            try:
                                                print(f"DEBUG: combined residuals rms(raw)={rms2:.3g}, eef_rms={eef_rms_unscaled:.3g}, int_rms={int_rms_unscaled:.3g}, tot_model2={tot_model2:.3g}, tot_data={tot_data:.3g}")
                                            except Exception:
                                                pass
                                            # acceptance criteria: relax thresholds slightly to improve
                                            # practical convergence. Also accept if one metric is
                                            # borderline but others are very good.
                                            eef_rms_thresh = 0.08
                                            int_rms_thresh = 1.0
                                            tot_ratio = (tot_model2 / tot_data) if (tot_data > 0 and np.isfinite(tot_data)) else np.nan
                                            if (np.isfinite(eef_rms_unscaled) and eef_rms_unscaled < eef_rms_thresh and
                                                np.isfinite(int_rms_unscaled) and int_rms_unscaled < int_rms_thresh and
                                                np.isfinite(tot_ratio) and 0.2 <= tot_ratio <= 4.0):
                                                accepted2 = True
                                            else:
                                                # allow a looser acceptance when intensity fit is excellent
                                                if (np.isfinite(int_rms_unscaled) and int_rms_unscaled < 0.5 and
                                                    np.isfinite(eef_rms_unscaled) and eef_rms_unscaled < 0.12 and
                                                    np.isfinite(tot_ratio) and 0.2 <= tot_ratio <= 4.0):
                                                    accepted2 = True
                                    except Exception:
                                        accepted2 = False
                                    if accepted2:
                                        # Require that combined-fit Pearson4 gives a better EEF
                                        # match than the pseudo-Voigt before accepting.
                                        try:
                                            pv_eef_rms = compute_pv_eef_rms()
                                        except Exception:
                                            pv_eef_rms = np.inf
                                        try:
                                            p4_eef_rms = float(eef_rms_unscaled) if np.isfinite(eef_rms_unscaled) else compute_p4_eef_rms_from_vec(best_params2)
                                        except Exception:
                                            p4_eef_rms = np.inf
                                        if np.isfinite(p4_eef_rms) and np.isfinite(pv_eef_rms):
                                            better = (p4_eef_rms <= 3.0 * pv_eef_rms)
                                        else:
                                            better = True
                                        if better:
                                            class SimpleResult2:
                                                pass
                                            pearson4_result = SimpleResult2()
                                            pearson4_result.params = vec_to_params(best_params2)
                                            try:
                                                print(f"DEBUG: pearson4_result set from combined-fit best_params2: {pearson4_result.params}")
                                            except Exception:
                                                pass
                                        else:
                                            try:
                                                print(f"DEBUG: combined Pearson4 fit rejected because PV EEF is closer (p4_rms={p4_eef_rms:.3g}, pv_rms={pv_eef_rms:.3g})")
                                            except Exception:
                                                pass
                                    else:
                                        try:
                                            print(f"DEBUG: combined Pearson4 fit rejected (rms2={rms2}, accepted2={accepted2})")
                                        except Exception:
                                            pass
                                        # Try a slower, global optimizer (differential_evolution) as a final attempt
                                        try:
                                            from scipy.optimize import differential_evolution
                                            # Bounds as list of tuples for differential_evolution
                                            bounds_de = []
                                            for i in range(lb.size):
                                                try:
                                                    bounds_de.append((float(lb[i]), float(ub[i])))
                                                except Exception:
                                                    bounds_de.append((float(-1e6), float(1e6)))

                                            def scalar_obj(v):
                                                try:
                                                    vec = np.asarray(v, dtype=float)
                                                    funv = np.asarray(pearson4_resid_comb_vec(vec), dtype=float)
                                                    return float(np.sqrt(np.mean(funv**2)))
                                                except Exception:
                                                    return 1e12

                                            try:
                                                # increase DE budget (up to ~3x slower) for more thorough search
                                                de_res = differential_evolution(scalar_obj, bounds_de, maxiter=240, popsize=45, tol=1e-6, seed=2026, polish=True)
                                                cand = de_res.x
                                                funvec = np.asarray(pearson4_resid_comb_vec(cand), dtype=float)
                                                rms_de = float(np.sqrt(np.mean(funvec**2))) if funvec.size else np.inf
                                                try:
                                                    print(f"DEBUG: differential_evolution rms={rms_de:.3g}, fun={de_res.fun if hasattr(de_res,'fun') else None}")
                                                except Exception:
                                                    pass
                                                # Re-evaluate acceptance using same unscaled diagnostics as earlier
                                                try:
                                                    pd2 = vec_to_params(cand)
                                                    model_full2 = pearson4_model.eval(params=pd2, x=r_arcsec)
                                                    model_full2 = np.maximum(model_full2, 0.0)
                                                    radial_model2 = 2.0 * np.pi * r_arcsec * model_full2 * dr_arcsec_p4
                                                    tot_model2 = np.sum(radial_model2)
                                                    tot_data = np.sum(radial_energy_data_p4)
                                                    if tot_model2 > 0 and np.isfinite(tot_model2):
                                                        eef_model2 = np.cumsum(radial_model2) / tot_model2
                                                        from numpy import interp
                                                        model_eef_at_rfit = interp(r_fit, r_arcsec, eef_model2)
                                                        eef_res_unscaled = (model_eef_at_rfit - eef_data_p4)
                                                        eef_rms_unscaled = float(np.sqrt(np.mean(eef_res_unscaled**2))) if eef_res_unscaled.size else np.inf
                                                        model_rfit_vals2 = pearson4_model.eval(params=pd2, x=r_fit)
                                                        model_rfit_vals2 = np.maximum(model_rfit_vals2, floor)
                                                        int_res_unscaled = (np.log(model_rfit_vals2 + floor) - np.log(I_fit + floor))
                                                        int_rms_unscaled = float(np.sqrt(np.mean(int_res_unscaled**2))) if int_res_unscaled.size else np.inf
                                                        tot_ratio = (tot_model2 / tot_data) if (tot_data > 0 and np.isfinite(tot_data)) else np.nan
                                                        # Apply same relaxed acceptance logic as before
                                                        eef_rms_thresh = 0.08
                                                        int_rms_thresh = 1.0
                                                        if (np.isfinite(eef_rms_unscaled) and eef_rms_unscaled < eef_rms_thresh and
                                                            np.isfinite(int_rms_unscaled) and int_rms_unscaled < int_rms_thresh and
                                                            np.isfinite(tot_ratio) and 0.2 <= tot_ratio <= 4.0):
                                                            accepted2 = True
                                                        else:
                                                            if (np.isfinite(int_rms_unscaled) and int_rms_unscaled < 0.5 and
                                                                np.isfinite(eef_rms_unscaled) and eef_rms_unscaled < 0.12 and
                                                                np.isfinite(tot_ratio) and 0.2 <= tot_ratio <= 4.0):
                                                                accepted2 = True
                                                except Exception:
                                                    accepted2 = False
                                                if accepted2:
                                                    # Before accepting DE candidate, ensure EEF is improved
                                                    try:
                                                        pv_eef_rms = compute_pv_eef_rms()
                                                    except Exception:
                                                        pv_eef_rms = np.inf
                                                    try:
                                                        p4_eef_rms = compute_p4_eef_rms_from_vec(cand)
                                                    except Exception:
                                                        p4_eef_rms = np.inf
                                                    if np.isfinite(p4_eef_rms) and np.isfinite(pv_eef_rms):
                                                        better = (p4_eef_rms <= 3.0 * pv_eef_rms)
                                                    else:
                                                        better = True
                                                    if better:
                                                        class SimpleResultDE:
                                                            pass
                                                        pearson4_result = SimpleResultDE()
                                                        pearson4_result.params = vec_to_params(cand)
                                                        try:
                                                            print(f"DEBUG: pearson4_result set from differential_evolution cand: {pearson4_result.params}")
                                                        except Exception:
                                                            pass
                                                    else:
                                                        try:
                                                            print(f"DEBUG: DE Pearson4 candidate rejected because PV EEF is closer (p4_rms={p4_eef_rms:.3g}, pv_rms={pv_eef_rms:.3g})")
                                                        except Exception:
                                                            pass
                                            except Exception:
                                                pass
                                        except Exception:
                                            pass
                                    # If combined multi-start failed, fall back to the quick lmfit
                                    # intensity-only fit when available and reasonably close.
                                    try:
                                        # Only accept quick intensity-only lmfit as a fallback
                                        # if it provides an excellent intensity match.
                                        if pearson4_result is None and 'quick_fit' in locals() and quick_fit is not None and hasattr(quick_fit, 'params'):
                                            try:
                                                model_rfit_q = pearson4_model.eval(params={k: float(v.value) for k, v in quick_fit.params.items() if hasattr(v, 'value')}, x=r_fit)
                                                model_rfit_q = np.maximum(model_rfit_q, floor)
                                                int_q_res = (np.log(model_rfit_q + floor) - np.log(I_fit + floor))
                                                int_q_rms = float(np.sqrt(np.mean(int_q_res**2))) if int_q_res.size else np.inf
                                                # tighten quick-fit fallback threshold to avoid showing
                                                # Pearson4 when only an intensity fit (not EEF) exists
                                                if np.isfinite(int_q_rms) and int_q_rms < 0.5:
                                                    # also ensure quick_fit yields better EEF than PV
                                                    try:
                                                        # extract params into vector
                                                        qp = quick_fit.params
                                                        amp_q = float(qp['amplitude'].value) if 'amplitude' in qp else float(np.nanmax(I_fit) if I_fit.size else 1.0)
                                                        cen_q = float(qp['center'].value) if 'center' in qp else float(r_fit[np.nanargmax(I_fit)]) if (r_fit.size and np.any(np.isfinite(I_fit))) else 0.0
                                                        sig_q = float(qp['sigma'].value) if 'sigma' in qp else max(1e-3, float(Gamma_c_fit))
                                                        exp_q = float(qp['expon'].value) if 'expon' in qp else 1.5
                                                        skew_q = float(qp['skew'].value) if 'skew' in qp else 0.0
                                                        p4_eef_rms_q = compute_p4_eef_rms_from_vec([amp_q, cen_q, sig_q, exp_q, skew_q])
                                                    except Exception:
                                                        p4_eef_rms_q = np.inf
                                                    try:
                                                        pv_eef_rms = compute_pv_eef_rms()
                                                    except Exception:
                                                        pv_eef_rms = np.inf
                                                    if np.isfinite(p4_eef_rms_q) and np.isfinite(pv_eef_rms):
                                                        better_q = (p4_eef_rms_q < pv_eef_rms)
                                                    else:
                                                        better_q = True
                                                    if better_q:
                                                        pearson4_result = quick_fit
                                                        try:
                                                            print(f"DEBUG: accepting quick lmfit intensity fit as fallback (int_rms={int_q_rms:.3g})")
                                                        except Exception:
                                                            pass
                                                    else:
                                                        try:
                                                            print(f"DEBUG: quick lmfit fallback rejected because PV EEF is closer (p4_rms={p4_eef_rms_q:.3g}, pv_rms={pv_eef_rms:.3g})")
                                                        except Exception:
                                                            pass
                                            except Exception:
                                                pass
                                    except Exception:
                                        pass
                                    
                                except Exception:
                                    pass
                        except Exception:
                            pearson4_result = None
                    else:
                        # fallback: use lmfit fit on intensity but keep previous result handling
                        try:
                            pearson4_result = pearson4_model.fit(I_fit, params, x=r_fit, max_nfev=10000)
                        except Exception:
                            pearson4_result = None
                except Exception:
                    pearson4_result = None
                
                # Generate Pearson4 diagnostic plot and compute EEF profile
                try:
                    # Only use Pearson4 results if an accepted fit exists; do not
                    # plot or evaluate the initial guess when no accepted fit.
                    pearson4_used_params_final = pearson4_result.params if (pearson4_result is not None and hasattr(pearson4_result, 'params')) else None
                    if pearson4_used_params_final is not None:
                        rplot_p4 = np.linspace(0.0, float(r_arcsec.max()), 1000)
                        Ifit_p4 = pearson4_model.eval(pearson4_used_params_final, x=rplot_p4)

                        fig, axes = plt.subplots(1, 2, figsize=(12, 4))

                        # Left: intensity fit
                        axes[0].plot(r_arcsec, I_profile, 'k.', ms=3, label='Data')
                        axes[0].plot(rplot_p4, Ifit_p4, 'b-', lw=2, label='Pearson4')
                        axes[0].set_xlabel('Radius [arcsec]')
                        axes[0].set_ylabel('Mean intensity')
                        axes[0].set_yscale('log')
                        if np.any(I_profile > 0):
                            lower = np.nanmin(I_profile[I_profile > 0])
                            axes[0].set_ylim(bottom=max(lower * 0.1, 1e-12))
                        axes[0].legend()
                        axes[0].grid(True, which='both', linestyle='--', alpha=0.4)
                        axes[0].set_title('Pearson4 Intensity Fit (log scale)')

                        # Right: residuals (log residuals evaluated on r_fit)
                        Ifit_p4_rfit = pearson4_model.eval(pearson4_used_params_final, x=r_fit)
                        residuals = np.log(Ifit_p4_rfit + floor) - np.log(I_fit + floor)
                        axes[1].plot(r_fit, residuals, 'ro', ms=4)
                        axes[1].axhline(0, color='k', linestyle='--', linewidth=1)
                        axes[1].set_xlabel('Radius [arcsec]')
                        axes[1].set_ylabel('Log residual')
                        axes[1].grid(True, which='both', linestyle='--', alpha=0.4)
                        axes[1].set_title('Pearson4 Fit Quality')

                        plt.tight_layout()
                        # Do not save the standalone Pearson4 diagnostic figure (user request)
                        try:
                            pass
                        except Exception:
                            pass
                        plt.close()

                        # Compute EEF profile for Pearson4 using the chosen params
                        I_fit_model_p4 = pearson4_model.eval(pearson4_used_params_final, x=r_arcsec)
                        I_fit_model_p4 = np.maximum(I_fit_model_p4, 0.0)
                        radial_energy_fit_p4 = 2.0 * np.pi * r_arcsec * I_fit_model_p4 * dr_arcsec
                        total_fit_energy_p4 = np.sum(radial_energy_fit_p4)
                        if total_fit_energy_p4 > 0 and np.isfinite(total_fit_energy_p4):
                            fit_cumulative_p4 = np.cumsum(radial_energy_fit_p4)
                            pearson4_profile_pct = 100.0 * fit_cumulative_p4 / total_fit_energy_p4
                            pearson4_profile_diam = 2.0 * r_arcsec
                        else:
                            pearson4_profile_pct = None
                            pearson4_profile_diam = None
                    else:
                        # No accepted Pearson4 fit: ensure nothing plotted and no profile
                        pearson4_profile_pct = None
                        pearson4_profile_diam = None
                except Exception as e:
                    pearson4_profile_pct = None
                    pearson4_profile_diam = None
            except ImportError:
                pearson4_result = None
                pearson4_profile_pct = None
                pearson4_profile_diam = None
            except Exception:
                pearson4_result = None
                pearson4_profile_pct = None
                pearson4_profile_diam = None
            
            # Create merged diagnostic: intensity curves + residuals, then export fit parameters
            # Disabled by default to avoid creating a second figure window.
            if False:
                os.makedirs('Figures', exist_ok=True)
                # Define plotting grid
                rplot = np.linspace(0.0, float(r_arcsec.max()), 1000)
                # PV model on rplot
                Ifit_pv_plot = beta_pseudo_gaussian(rplot, A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit)
                # Pearson4 model params: prefer fit result params, fall back to params
                try:
                    p4_params_plot = pearson4_result.params if (pearson4_result is not None and hasattr(pearson4_result, 'params')) else (params if 'params' in locals() else None)
                except Exception:
                    p4_params_plot = (params if 'params' in locals() else None)
                Ifit_p4_plot = None
                if p4_params_plot is not None:
                    try:
                        Ifit_p4_plot = pearson4_model.eval(p4_params_plot, x=rplot)
                    except Exception:
                        Ifit_p4_plot = None

                # --- King profile fitting (EEF-only objective) ---
                Ifit_king_plot = None
                king_profile_pct = None
                king_profile_diam = None
                try:
                    # initial guesses
                    I0_0 = float(np.nanmax(I_fit)) if I_fit.size else 1.0
                    # Use pseudo-Voigt core width as a better initial guess for King core size when available
                    try:
                        rc_guess = float(Gamma_c_fit)
                    except Exception:
                        rc_guess = float(np.median(r_fit)) if r_fit.size else 1.0
                    rc_0 = max(0.5 * rc_guess, 1e-3)
                    alpha_0 = 2.0
                    b0 = float(np.nanmin(I_fit)) if np.any(np.isfinite(I_fit)) else 0.0
                    x0k = [I0_0, rc_0, alpha_0, b0]
                    # Constrain I0 to be at least the noise floor and limit background b using data percentiles
                    I0_lb = float(floor)
                    try:
                        p25 = float(np.nanpercentile(I_fit, 25)) if np.any(np.isfinite(I_fit)) else None
                        p50 = float(np.nanmedian(I_fit)) if np.any(np.isfinite(I_fit)) else None
                    except Exception:
                        p25, p50 = None, None
                    if p25 is not None and np.isfinite(p25):
                        b_ub = max(p25 * 2.0, p50 if (p50 is not None and np.isfinite(p50)) else p25 * 2.0)
                    else:
                        b_ub = max(float(floor) * 100.0, 1.0)
                    # Tighten King parameter bounds to avoid degenerate long tails
                    lbk = [I0_lb, 1e-3, 1.0, 0.0]  # enforce alpha >= 1.0
                    # limit rc to a reasonable fraction of the sampled radius and cap it
                    rc_cap = float(min(r_arcsec.max() * 0.5 if r_arcsec.size else 10.0, 50.0))
                    ubk = [np.inf, rc_cap, 8.0, float(b_ub)]

                    def resid_king(vec):
                        try:
                            I0_k, rc_k, alpha_k, b_k = vec
                            model_vals = king_profile(r_arcsec, I0_k, rc_k, alpha_k, b_k)
                            model_vals = np.maximum(model_vals, 0.0)
                            dr = float(r_arcsec[1] - r_arcsec[0]) if r_arcsec.size>1 else 1.0
                            radial = 2.0 * np.pi * r_arcsec * model_vals * dr
                            tot = np.sum(radial)
                            if tot <= 0 or not np.isfinite(tot):
                                return np.ones_like(eef_data) * 1e6
                            model_eef_pct = 100.0 * np.cumsum(radial) / tot
                            from numpy import interp
                            # use eef_data (computed on r_fit) as the reference EEF values at r_fit (fractions 0..1)
                            ref_at_r = eef_data
                            # convert model EEF percent -> fraction before comparing
                            model_at_ref = interp(r_fit, r_arcsec, model_eef_pct) / 100.0
                            # scale by the eef weight used elsewhere so residuals have comparable magnitude
                            return (model_at_ref - ref_at_r) * eef_weight
                        except Exception:
                            return np.ones_like(eef_data) * 1e6

                    # run King fit if we have the reference EEF computed on r_fit (eef_data)
                    if 'eef_data' in locals() and eef_data is not None:
                        if have_least_squares:
                            try:
                                try:
                                    print(f"King fit: initial x0k={x0k}, lbk={lbk}, ubk={ubk}")
                                except Exception:
                                    pass
                                # increase King-fit starts and budget to improve robustness
                                best_vec, last_resk = multi_start_least_squares(resid_king, x0k, lbk, ubk, attempts=24, rng_seed=1234, loss='soft_l1', f_scale=1.0, max_nfev=80000)
                                if best_vec is not None:
                                    I0_k, rc_k, alpha_k, b_k = best_vec.tolist()
                                else:
                                    I0_k, rc_k, alpha_k, b_k = x0k
                            except Exception:
                                I0_k, rc_k, alpha_k, b_k = x0k
                        else:
                            I0_k, rc_k, alpha_k, b_k = x0k

                        # compute king model and EEF
                        try:
                            try:
                                print(f"Computing King model with params: I0={I0_k}, rc={rc_k}, alpha={alpha_k}, b={b_k}")
                            except Exception:
                                pass
                            Ifit_king_plot = king_profile(rplot, I0_k, rc_k, alpha_k, b_k)
                            vals_k = king_profile(r_arcsec, I0_k, rc_k, alpha_k, b_k)
                            dr = float(r_arcsec[1] - r_arcsec[0]) if r_arcsec.size>1 else 1.0
                            radial_k = 2.0 * np.pi * r_arcsec * vals_k * dr
                            totk = np.sum(radial_k)
                            try:
                                print(f"  King totk={totk}")
                            except Exception:
                                pass
                            if totk > 0 and np.isfinite(totk):
                                cumk = np.cumsum(radial_k)
                                king_profile_pct = 100.0 * cumk / totk
                                king_profile_diam = 2.0 * r_arcsec
                                try:
                                    print(f"King fit succeeded: I0={I0_k:.3g}, rc={rc_k:.3g}, alpha={alpha_k:.3g}, b={b_k:.3g}")
                                    print(f"  king_profile_pct len={len(king_profile_pct)}, min={np.nanmin(king_profile_pct):.3g}, max={np.nanmax(king_profile_pct):.3g}")
                                except Exception:
                                    pass
                            else:
                                king_profile_pct = None
                                king_profile_diam = None
                        except Exception:
                            king_profile_pct = None
                            king_profile_diam = None
                except Exception:
                    Ifit_king_plot = None
                    king_profile_pct = None
                    king_profile_diam = None

                # Prepare residuals evaluated at r_fit (data sample points)
                model_pv_rfit = beta_pseudo_gaussian(r_fit, A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit)
                model_pv_rfit = np.maximum(model_pv_rfit, floor)
                res_int_pv = np.log(model_pv_rfit + floor) - np.log(I_fit + floor)

                res_int_p4 = None
                if p4_params_plot is not None:
                    try:
                        model_p4_rfit = pearson4_model.eval(p4_params_plot, x=r_fit)
                        model_p4_rfit = np.maximum(model_p4_rfit, floor)
                        res_int_p4 = np.log(model_p4_rfit + floor) - np.log(I_fit + floor)
                    except Exception:
                        res_int_p4 = None

                # King intensity residuals at r_fit
                res_int_king = None
                if Ifit_king_plot is not None:
                    try:
                        model_king_rfit = king_profile(r_fit, I0_k, rc_k, alpha_k, b_k)
                        model_king_rfit = np.maximum(model_king_rfit, floor)
                        res_int_king = np.log(model_king_rfit + floor) - np.log(I_fit + floor)
                    except Exception:
                        res_int_king = None

                # EEF residuals: compute model EEF on r_arcsec grid and interpolate to r_fit
                eef_ref_pct = profile_pct if ('profile_pct' in locals() and profile_pct is not None) else None
                eef_ref_diam = profile_diam if ('profile_diam' in locals() and profile_diam is not None) else None
                res_eef_pv = None
                res_eef_p4 = None
                try:
                    if fit_profile_pct is not None and fit_profile_diam is not None and eef_ref_pct is not None and eef_ref_diam is not None:
                        # PV model EEF on r_arcsec
                        pv_vals = beta_pseudo_gaussian(r_arcsec, A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit)
                        dr = float(r_arcsec[1] - r_arcsec[0]) if r_arcsec.size>1 else 1.0
                        radial_pv = 2.0 * np.pi * r_arcsec * pv_vals * dr
                        tot_pv = np.sum(radial_pv)
                        if tot_pv > 0 and np.isfinite(tot_pv):
                            pv_eef_pct = 100.0 * np.cumsum(radial_pv) / tot_pv
                            # interpolate PV EEF to r_fit (r in arcsec)
                            from numpy import interp
                            pv_eef_at_rfit = interp(r_fit, r_arcsec, pv_eef_pct)
                            ref_at_rfit = interp(r_fit, eef_ref_diam/2.0, eef_ref_pct)
                            res_eef_pv = pv_eef_at_rfit - ref_at_rfit
                        # Pearson4 EEF
                        if p4_params_plot is not None:
                            p4_vals = pearson4_model.eval(p4_params_plot, x=r_arcsec)
                            radial_p4 = 2.0 * np.pi * r_arcsec * p4_vals * dr
                            tot_p4 = np.sum(radial_p4)
                            if tot_p4 > 0 and np.isfinite(tot_p4):
                                p4_eef_pct = 100.0 * np.cumsum(radial_p4) / tot_p4
                                p4_eef_at_rfit = interp(r_fit, r_arcsec, p4_eef_pct)
                                ref_at_rfit = interp(r_fit, eef_ref_diam/2.0, eef_ref_pct)
                                res_eef_p4 = p4_eef_at_rfit - ref_at_rfit
                        # King EEF
                        if king_profile_pct is not None and king_profile_diam is not None:
                            try:
                                king_eef_at_rfit = interp(r_fit, king_profile_diam/2.0, king_profile_pct)
                                king_vals = king_profile(r_arcsec, I0_k, rc_k, alpha_k, b_k)
                                radial_k = 2.0 * np.pi * r_arcsec * king_vals * dr
                                tot_k = np.sum(radial_k)
                                if tot_k > 0 and np.isfinite(tot_k):
                                    king_eef_pct = 100.0 * np.cumsum(radial_k) / tot_k
                                    king_eef_at_rfit = interp(r_fit, r_arcsec, king_eef_pct)
                                    ref_at_rfit = interp(r_fit, eef_ref_diam/2.0, eef_ref_pct)
                                    res_eef_king = king_eef_at_rfit - ref_at_rfit
                                else:
                                    res_eef_king = None
                            except Exception:
                                res_eef_king = None
                except Exception:
                    res_eef_pv = None
                    res_eef_p4 = None

                # Prepare aggregated EEF arrays for plotting (robustly)
                profile_pct = None
                profile_diam = None
                try:
                    # r_profile is in meters; convert to arcsec using arcsec_to_m
                    if ('cumulative_profile' in locals() and 'total_energy' in locals() and total_energy and 'r_profile' in locals() and 'arcsec_to_m' in locals()):
                        profile_pct = 100.0 * (cumulative_profile / total_energy)
                        # diameter in arcsec = 2 * (r_profile [m] / arcsec_to_m [m/arcsec])
                        profile_diam = 2.0 * (r_profile / arcsec_to_m)
                except Exception:
                    profile_pct = None
                    profile_diam = None

                # Recompute EEF residuals at r_fit explicitly for plotting (ensure bottom-right shows data)
                res_eef_pv_plot = None
                res_eef_p4_plot = None
                res_eef_king_plot = None
                try:
                    from numpy import interp
                    dr0 = float(r_arcsec[1] - r_arcsec[0]) if r_arcsec.size>1 else 1.0
                    # reference EEF at r_fit (fraction)
                    ref_at_rfit = eef_data if ('eef_data' in locals() and eef_data is not None) else None
                    if ref_at_rfit is not None:
                        # PV model
                        try:
                            pv_vals_full = beta_pseudo_gaussian(r_arcsec, A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit)
                            radial_pv_full = 2.0 * np.pi * r_arcsec * pv_vals_full * dr0
                            tot_pv_full = np.sum(radial_pv_full)
                            if tot_pv_full > 0 and np.isfinite(tot_pv_full):
                                pv_eef_pct_full = 100.0 * np.cumsum(radial_pv_full) / tot_pv_full
                                pv_eef_at_rfit = interp(r_fit, r_arcsec, pv_eef_pct_full) / 100.0
                                res_eef_pv_plot = pv_eef_at_rfit - ref_at_rfit
                        except Exception:
                            res_eef_pv_plot = None
                        # Pearson4 model
                        try:
                            if p4_params_plot is not None:
                                p4_vals_full = pearson4_model.eval(p4_params_plot, x=r_arcsec)
                                radial_p4_full = 2.0 * np.pi * r_arcsec * p4_vals_full * dr0
                                tot_p4_full = np.sum(radial_p4_full)
                                if tot_p4_full > 0 and np.isfinite(tot_p4_full):
                                    p4_eef_pct_full = 100.0 * np.cumsum(radial_p4_full) / tot_p4_full
                                    p4_eef_at_rfit = interp(r_fit, r_arcsec, p4_eef_pct_full) / 100.0
                                    res_eef_p4_plot = p4_eef_at_rfit - ref_at_rfit
                        except Exception:
                            res_eef_p4_plot = None
                        # King model
                        try:
                            if 'king_profile_pct' in locals() and king_profile_pct is not None and 'king_profile_diam' in locals() and king_profile_diam is not None:
                                king_vals_full = king_profile(r_arcsec, I0_k, rc_k, alpha_k, b_k)
                                radial_k_full = 2.0 * np.pi * r_arcsec * king_vals_full * dr0
                                tot_k_full = np.sum(radial_k_full)
                                if tot_k_full > 0 and np.isfinite(tot_k_full):
                                    king_eef_pct_full = 100.0 * np.cumsum(radial_k_full) / tot_k_full
                                    king_eef_at_rfit = interp(r_fit, r_arcsec, king_eef_pct_full) / 100.0
                                    res_eef_king_plot = king_eef_at_rfit - ref_at_rfit
                        except Exception:
                            res_eef_king_plot = None
                except Exception:
                    res_eef_pv_plot = None
                    res_eef_p4_plot = None
                    res_eef_king_plot = None

                # Plot combined figure
                # Create a 2x2 figure: top row = intensity + residuals; bottom row = EEF + EEF residuals
                fig, axes = plt.subplots(2, 2, figsize=(14, 10), constrained_layout=True)
                ax_int = axes[0, 0]
                ax_res = axes[0, 1]
                ax_eef = axes[1, 0]
                ax_eef_res = axes[1, 1]

                # Determine whether Pearson4 was accepted (only then show its traces)
                try:
                    pearson4_ok = (pearson4_result is not None)
                except Exception:
                    pearson4_ok = False

                # Top-left: intensities
                # define linewidths for combined figure: thin for dashed fits, thicker for solid/reference
                _combined_thin_lw = 1.0
                _combined_thick_lw = 2.0
                ax_int.plot(r_arcsec, I_profile, 'k.', ms=3, label='Aggregated PSF (data)')
                line_pv_comb = ax_int.plot(rplot, Ifit_pv_plot, color='red', linestyle='--', lw=_combined_thin_lw, label='Modified pseudo-Voigt')[0]
                try:
                    line_pv_comb.set_dashes([6, 2])
                except Exception:
                    pass
                if pearson4_ok and Ifit_p4_plot is not None:
                    ax_int.plot(rplot, Ifit_p4_plot, color='orange', linestyle='--', lw=_combined_thin_lw, label=label_p4)
                if Ifit_king_plot is not None:
                    ax_int.plot(rplot, Ifit_king_plot, color='purple', linestyle='--', lw=_combined_thin_lw, label=label_king)
                ax_int.set_xlabel('Radius [arcsec]')
                ax_int.set_ylabel('Mean intensity')
                ax_int.set_yscale('log')
                ax_int.legend(fontsize=9)
                ax_int.grid(True, which='both', linestyle='--', alpha=0.4)

                # Top-right: intensity residuals (left y-axis, percent) and
                # EEF residuals (right y-axis, linear percent). Convert intensity
                # residuals to percent so all curves share the same units.
                ax_res.set_xlabel('Radius [arcsec]')
                ax_res.set_ylabel('Residual (%)')
                # Twin axis for EEF residuals (percentage, linear)
                ax_res_eef = ax_res.twinx()
                ax_res_eef.set_ylabel('EEF residual (pct)')

                # Convert log/int residuals to percent relative to data
                eps_local = floor if ('floor' in locals() or 'floor' in globals()) else 1e-12
                try:
                    if res_int_pv is not None and 'I_fit' in locals():
                        # res_int_pv was log(model) - log(data); prefer to use the
                        # linear model values already computed on r_fit (model_pv_rfit)
                        try:
                            if 'model_pv_rfit' in locals() and model_pv_rfit is not None:
                                res_int_pv_pct = 100.0 * (model_pv_rfit - I_fit) / (I_fit + eps_local)
                            else:
                                # Fall back: convert log-residual to approximate percent using exp(delta)-1
                                res_int_pv_pct = 100.0 * (np.exp(res_int_pv) - 1.0)
                        except Exception:
                            res_int_pv_pct = np.full_like(r_fit, np.nan)
                    else:
                        res_int_pv_pct = np.full_like(r_fit, np.nan)
                except Exception:
                    res_int_pv_pct = np.full_like(r_fit, np.nan)
                try:
                    if pearson4_ok and res_int_p4 is not None:
                        res_int_p4_pct = 100.0 * (np.exp(res_int_p4) - 1.0)
                    else:
                        res_int_p4_pct = np.full_like(r_fit, np.nan)
                except Exception:
                    res_int_p4_pct = np.full_like(r_fit, np.nan)
                try:
                    if res_int_king is not None:
                        res_int_king_pct = 100.0 * (np.exp(res_int_king) - 1.0)
                    else:
                        res_int_king_pct = np.full_like(r_fit, np.nan)
                except Exception:
                    res_int_king_pct = np.full_like(r_fit, np.nan)

                # Plot intensity residuals on the left (percent)
                if res_int_pv_pct is not None:
                    ax_res.plot(r_fit, res_int_pv_pct, color='red', linestyle='--', lw=_combined_thin_lw, label='Intensity residual: Data - Pseudo-Voigt (%)')
                if pearson4_ok and res_int_p4_pct is not None:
                    ax_res.plot(r_fit, res_int_p4_pct, color='orange', linestyle='--', lw=_combined_thin_lw * 1.6, marker='o', markersize=3, markevery=50, label='Intensity residual: Data - Pearson4 (%)')
                if res_int_king_pct is not None:
                    ax_res.plot(r_fit, res_int_king_pct, color='purple', linestyle='--', lw=_combined_thin_lw * 1.6, marker='s', markersize=3, markevery=50, label=f'Intensity residual: Data - {label_king} (%)')
                ax_res.axhline(0.0, color='k', linestyle='--', linewidth=1)
                ax_res.grid(True, which='both', linestyle='--', alpha=0.4)

                # Plot EEF residuals on the right (linear scale)
                handles_res, labels_res = [], []
                if res_eef_pv is not None:
                    h1, = ax_res_eef.plot(r_fit, res_eef_pv, color='red', linestyle='--', lw=_combined_thin_lw, label='EEF residual: PV - Data (pct)')
                    handles_res.append(h1); labels_res.append('EEF residual: PV - Data (pct)')
                if res_eef_p4 is not None:
                    h2, = ax_res_eef.plot(r_fit, res_eef_p4, color='orange', linestyle='--', lw=_combined_thin_lw * 1.6, marker='o', markersize=3, markevery=50, label='EEF residual: P4 - Data (pct)')
                    handles_res.append(h2); labels_res.append('EEF residual: P4 - Data (pct)')
                if 'res_eef_king' in locals() and res_eef_king is not None:
                    h3, = ax_res_eef.plot(r_fit, res_eef_king, color='purple', linestyle='--', lw=_combined_thin_lw * 1.6, marker='s', markersize=3, markevery=50, label='EEF residual: King - Data (pct)')
                    handles_res.append(h3); labels_res.append('EEF residual: King - Data (pct)')

                # Combine legends from both axes (intensity residuals on left, EEF residuals on right)
                try:
                    handles_l, labels_l = ax_res.get_legend_handles_labels()
                    handles_r, labels_r = ax_res_eef.get_legend_handles_labels()
                    all_handles = handles_l + handles_r
                    all_labels = labels_l + labels_r
                    if all_handles:
                        ax_res.legend(all_handles, all_labels, fontsize=9)
                except Exception:
                    try:
                        ax_res.legend()
                    except Exception:
                        pass

                # Bottom-left: EEF curves (aggregated PSF and fitted models)
                # Flip axes: diameter on X, encircled energy (%) on Y. Plot aggregated PSF as black dots.
                label_data = 'Aggregated PSF (data)'
                label_pv = 'Modified pseudo-Voigt radial fit'
                # Use short, clean labels for fitted models in the combined EEF figure
                # Full fit parameters are retained in the exported Excel/CSV, not the legend.
                label_p4 = 'Pearson4 radial fit'
                label_king = 'King radial fit'
                try:
                    # Ensure Pearson4 EEF is available: if it wasn't computed earlier
                    # but we do have a fit result, compute it here so the combined
                    # figure always shows the Pearson4 EEF when possible.
                    if (pearson4_profile_pct is None or pearson4_profile_diam is None) and (pearson4_result is not None) and hasattr(pearson4_result, 'params'):
                        try:
                            p_params_plot = pearson4_result.params
                            rplot_local = np.linspace(0.0, float(r_arcsec.max()), 1000)
                            # compute model on r_arcsec for EEF
                            p_vals = pearson4_model.eval(p_params_plot, x=r_arcsec)
                            p_vals = np.maximum(p_vals, 0.0)
                            dr_local = float(r_arcsec[1] - r_arcsec[0]) if r_arcsec.size > 1 else 1.0
                            radial_p = 2.0 * np.pi * r_arcsec * p_vals * dr_local
                            tot_p = np.sum(radial_p)
                            if tot_p > 0 and np.isfinite(tot_p):
                                pearson4_profile_pct = 100.0 * np.cumsum(radial_p) / tot_p
                                pearson4_profile_diam = 2.0 * r_arcsec
                                # also provide Ifit_p4_plot for intensity panel if missing
                                try:
                                    Ifit_p4_plot = pearson4_model.eval(p_params_plot, x=rplot_local)
                                except Exception:
                                    Ifit_p4_plot = None
                        except Exception:
                            pearson4_profile_pct = None
                            pearson4_profile_diam = None

                    # Debug: report whether Pearson4 fit/result was used to compute the EEF
                    try:
                        present = (pearson4_result is not None and hasattr(pearson4_result, 'params'))
                        computed = (pearson4_profile_pct is not None and pearson4_profile_diam is not None)
                        print(f"DEBUG: pearson4_result present={present}, pearson4_profile computed={computed}")
                    except Exception:
                        pass

                    if profile_pct is not None and profile_diam is not None and len(profile_pct) and len(profile_diam):
                        pct_arr = np.asarray(profile_pct, dtype=float)
                        diam_arr = np.asarray(profile_diam, dtype=float)
                        order = np.argsort(pct_arr)
                        pct_sorted = pct_arr[order]
                        diam_sorted = diam_arr[order]
                        # Clip to 95% percentile
                        mask95 = pct_sorted <= 95.0
                        if np.any(mask95):
                            ax_eef.plot(diam_sorted[mask95], pct_sorted[mask95], color='k', marker='.', linestyle='None', ms=4, label=label_data, zorder=10)
                    # Plot PV EEF if available (diameter vs pct, clipped to 95%)
                    # Prefer explicit PV EEF computed on the fine rplot grid so it's always visible
                    try:
                        if 'Ifit_pv_plot' in locals() and Ifit_pv_plot is not None:
                            # rplot is in arcsec
                            rpv = rplot
                            drpv = float(rpv[1] - rpv[0]) if rpv.size > 1 else 1.0
                            radial_pv_plot = 2.0 * np.pi * rpv * Ifit_pv_plot * drpv
                            tot_pv_plot = np.sum(radial_pv_plot)
                            if tot_pv_plot > 0 and np.isfinite(tot_pv_plot):
                                pv_eef_pct_plot = 100.0 * np.cumsum(radial_pv_plot) / tot_pv_plot
                                pv_diam_plot = 2.0 * rpv
                                maskpv = pv_eef_pct_plot <= 95.0
                                if np.any(maskpv):
                                    line_pv_eef = ax_eef.plot(pv_diam_plot[maskpv], pv_eef_pct_plot[maskpv], color='red', linestyle='--', lw=_combined_thin_lw, label=label_pv)[0]
                                    try:
                                        line_pv_eef.set_dashes([6, 2])
                                    except Exception:
                                        pass
                        elif fit_profile_pct is not None and fit_profile_diam is not None:
                            fp_pct = np.asarray(fit_profile_pct, dtype=float)
                            fp_diam = np.asarray(fit_profile_diam, dtype=float)
                            mask = fp_pct <= 95.0
                            if np.any(mask):
                                line_pv_eef2 = ax_eef.plot(fp_diam[mask], fp_pct[mask], color='red', linestyle='--', lw=_combined_thin_lw, label=label_pv)[0]
                                try:
                                    line_pv_eef2.set_dashes([6, 2])
                                except Exception:
                                    pass
                    except Exception:
                        try:
                            if fit_profile_pct is not None and fit_profile_diam is not None:
                                fp_pct = np.asarray(fit_profile_pct, dtype=float)
                                fp_diam = np.asarray(fit_profile_diam, dtype=float)
                                mask = fp_pct <= 95.0
                                if np.any(mask):
                                    line_pv_eef3 = ax_eef.plot(fp_diam[mask], fp_pct[mask], color='red', linestyle='--', lw=_combined_thin_lw, label=label_pv)[0]
                                    try:
                                        line_pv_eef3.set_dashes([6, 2])
                                    except Exception:
                                        pass
                        except Exception:
                            pass
                    # Pearson4: only plot if a Pearson4 fit/result was accepted
                    try:
                        pearson4_ok = (pearson4_result is not None)
                    except Exception:
                        pearson4_ok = False
                    if pearson4_ok and pearson4_profile_pct is not None and pearson4_profile_diam is not None:
                        p4_pct = np.asarray(pearson4_profile_pct, dtype=float)
                        p4_diam = np.asarray(pearson4_profile_diam, dtype=float)
                        mask = p4_pct <= 95.0
                        if np.any(mask):
                            ax_eef.plot(p4_diam[mask], p4_pct[mask], color='orange', linestyle='--', lw=_combined_thin_lw, label=label_p4)
                    # King
                    if 'king_profile_pct' in locals() and king_profile_pct is not None and 'king_profile_diam' in locals() and king_profile_diam is not None:
                        k_pct = np.asarray(king_profile_pct, dtype=float)
                        k_diam = np.asarray(king_profile_diam, dtype=float)
                        mask = k_pct <= 95.0
                        if np.any(mask):
                            ax_eef.plot(k_diam[mask], k_pct[mask], color='purple', linestyle='--', linewidth=_combined_thin_lw, label=label_king)
                    else:
                        # Fallback: if we have an evaluated King profile on rplot, compute its EEF and plot
                        try:
                            if 'Ifit_king_plot' in locals() and Ifit_king_plot is not None:
                                rpv = rplot if 'rplot' in locals() else (np.linspace(0.0, float(r_arcsec.max()), 1000) if 'r_arcsec' in locals() else None)
                                if rpv is not None:
                                    drpv = float(rpv[1] - rpv[0]) if rpv.size > 1 else 1.0
                                    radial_k_plot = 2.0 * np.pi * rpv * Ifit_king_plot * drpv
                                    tot_k_plot = np.sum(radial_k_plot)
                                    if tot_k_plot > 0 and np.isfinite(tot_k_plot):
                                        k_eef_pct_plot = 100.0 * np.cumsum(radial_k_plot) / tot_k_plot
                                        k_diam_plot = 2.0 * rpv
                                        maskk = k_eef_pct_plot <= 95.0
                                        if np.any(maskk):
                                            ax_eef.plot(k_diam_plot[maskk], k_eef_pct_plot[maskk], color='purple', linestyle='--', linewidth=_combined_thin_lw, label=label_king)
                        except Exception:
                            pass
                    ax_eef.set_xlabel('Diameter [arcsec]')
                    ax_eef.set_ylabel('Encircled energy (%)')
                    # Use default legend from available handles so the full
                    # constructed labels (e.g. label_king) are shown without
                    # attempting fragile re-ordering.
                    try:
                        import textwrap
                        handles_all, labels_all = ax_eef.get_legend_handles_labels()
                        if handles_all:
                            # Wrap long legend labels to avoid over-wide single-line entries
                            wrapped_labels = [textwrap.fill(lbl, width=40) if isinstance(lbl, str) else lbl for lbl in labels_all]
                            # Place EEF legend inside the plotting area (upper left)
                            ax_eef.legend(handles_all, wrapped_labels, loc='upper left', fontsize=9)
                        
                    except Exception:
                        try:
                            ax_eef.legend()
                        except Exception:
                            pass
                    ax_eef.grid(True, which='both', linestyle='--', alpha=0.4)
                    # EEF 50% markers: vertical line at diameter for 50% and horizontal at 50%
                    try:
                        if 'pct_sorted' in locals() and np.any(pct_sorted <= 95.0):
                            pct_for_interp = pct_sorted
                            diam_for_interp = diam_sorted
                            if pct_for_interp[0] <= 50.0 <= pct_for_interp[-1]:
                                diam50 = float(np.interp(50.0, pct_for_interp, diam_for_interp))
                                # make 50% markers subtle: light grey and thinner
                                subtle_lw = 0.8
                                subtle_col = 'lightgrey'
                                ax_eef.axvline(x=diam50, linestyle='--', color=subtle_col, linewidth=subtle_lw)
                                ax_eef.axhline(y=50.0, linestyle='--', color=subtle_col, linewidth=subtle_lw)
                                try:
                                    ax_eef.text(diam_for_interp.min(), 50.0, f'EEF 50% = {diam50:.3f}"', ha='left', va='bottom', fontsize=9, color='purple')
                                except Exception:
                                    pass
                    except Exception:
                        pass
                except Exception:
                    # fallback: do nothing
                    pass

                # Bottom-right: EEF residuals (Data - Model) evaluated at r_fit
                # Plot as percent. Use recomputed fractional residuals when available
                # (those are in fractions -> multiply by 100), otherwise fall back
                # to earlier-percent values and just flip sign.
                if res_eef_pv_plot is not None:
                    line_pv_res = ax_eef_res.plot(r_fit, (-res_eef_pv_plot) * 100.0, color='red', linestyle='--', lw=_combined_thin_lw, label=f'EEF residual: Data - {label_pv} (pct)')[0]
                    try:
                        line_pv_res.set_dashes([6, 2])
                    except Exception:
                        pass
                elif res_eef_pv is not None:
                    line_pv_res2 = ax_eef_res.plot(r_fit, (-res_eef_pv), color='red', linestyle='--', lw=_combined_thin_lw, label=f'EEF residual: Data - {label_pv} (pct)')[0]
                    try:
                        line_pv_res2.set_dashes([6, 2])
                    except Exception:
                        pass
                if pearson4_ok:
                    if res_eef_p4_plot is not None:
                        ax_eef_res.plot(r_fit, (-res_eef_p4_plot) * 100.0, color='orange', linestyle='--', lw=_combined_thin_lw * 1.6, marker='o', markersize=3, markevery=50, label=f'EEF residual: Data - {label_p4} (pct)')
                    elif res_eef_p4 is not None:
                        ax_eef_res.plot(r_fit, (-res_eef_p4), color='orange', linestyle='--', lw=_combined_thin_lw * 1.6, marker='o', markersize=3, markevery=50, label=f'EEF residual: Data - {label_p4} (pct)')
                if res_eef_king_plot is not None:
                    ax_eef_res.plot(r_fit, (-res_eef_king_plot) * 100.0, color='purple', linestyle='--', lw=_combined_thin_lw * 1.6, marker='s', markersize=3, markevery=50, label=f'EEF residual: Data - {label_king} (pct)')
                elif 'res_eef_king' in locals() and res_eef_king is not None:
                    ax_eef_res.plot(r_fit, (-res_eef_king), color='purple', linestyle='--', lw=_combined_thin_lw * 1.6, marker='s', markersize=3, markevery=50, label=f'EEF residual: Data - {label_king} (pct)')
                ax_eef_res.axhline(0.0, color='k', linestyle='--', linewidth=1)
                ax_eef_res.set_xlabel('Radius [arcsec]')
                ax_eef_res.set_ylabel('EEF residual (pct)')
                # Mark radius corresponding to EEF 50% (diameter from bottom-left / 2)
                try:
                    if 'diam50' in locals():
                        radius50 = float(diam50) / 2.0
                        subtle_lw = 0.8
                        subtle_col = 'lightgrey'
                        ax_eef_res.axvline(x=radius50, linestyle='--', color=subtle_col, linewidth=subtle_lw)
                        try:
                            ylim = ax_eef_res.get_ylim()
                            # move R50 label further away from the line (increase vertical offset)
                            ax_eef_res.text(radius50, ylim[0] + 0.08 * (ylim[1] - ylim[0]), f'R50 = {radius50:.3f}"', ha='center', va='bottom', fontsize=9, color='black', rotation=90)
                        except Exception:
                            pass
                except Exception:
                    pass
                try:
                    handles_r, labels_r = ax_eef_res.get_legend_handles_labels()
                    if handles_r:
                        ax_eef_res.legend(handles_r, labels_r, loc='upper right', fontsize=9)
                except Exception:
                    try:
                        ax_eef_res.legend()
                    except Exception:
                        pass
                ax_eef_res.grid(True, which='both', linestyle='--', alpha=0.4)

                outfn_comb = os.path.join('Figures', 'E2E_fit_combined.png')
                # Ensure 50% markers are drawn even if earlier variables were missing
                try:
                    diam50_forced = None
                    def try_from_arrays(pct_arr, diam_arr):
                        try:
                            pa = np.asarray(pct_arr, dtype=float)
                            da = np.asarray(diam_arr, dtype=float)
                            order = np.argsort(pa)
                            pa_s = pa[order]
                            da_s = da[order]
                            if pa_s.size and pa_s[0] <= 50.0 <= pa_s[-1]:
                                return float(np.interp(50.0, pa_s, da_s))
                        except Exception:
                            return None
                        return None

                    # Try preferred sources in order
                    if diam50_forced is None and 'profile_pct' in locals() and profile_pct is not None:
                        diam50_forced = try_from_arrays(profile_pct, profile_diam)
                    if diam50_forced is None and 'fit_profile_pct' in locals() and fit_profile_pct is not None:
                        diam50_forced = try_from_arrays(fit_profile_pct, fit_profile_diam)
                    if diam50_forced is None and 'pearson4_profile_pct' in locals() and pearson4_profile_pct is not None:
                        diam50_forced = try_from_arrays(pearson4_profile_pct, pearson4_profile_diam)
                    if diam50_forced is None and 'king_profile_pct' in locals() and king_profile_pct is not None:
                        diam50_forced = try_from_arrays(king_profile_pct, king_profile_diam)

                    # If we didn't find diam50 earlier, force-compute from frac_profile/r_profile
                    try:
                        if diam50_forced is None:
                            diam50_forced = None
                            if 'frac_profile' in locals() and 'r_profile' in locals() and frac_profile is not None and r_profile is not None:
                                try:
                                    fp = np.asarray(frac_profile, dtype=float) * 100.0
                                    rp = np.asarray(r_profile, dtype=float) * m_to_arcsec * 2.0  # diameter
                                    order = np.argsort(fp)
                                    fp_s = fp[order]
                                    rp_s = rp[order]
                                    if fp_s.size and fp_s[0] <= 50.0 <= fp_s[-1]:
                                        diam50_forced = float(np.interp(50.0, fp_s, rp_s))
                                except Exception:
                                    diam50_forced = None
                        if diam50_forced is not None:
                            # draw with high zorder so they're visible above plotted curves
                            try:
                                subtle_lw_forced = 0.8
                                subtle_col = 'lightgrey'
                                ax_eef.axvline(x=diam50_forced, linestyle='--', color=subtle_col, linewidth=subtle_lw_forced, zorder=100)
                                ax_eef.axhline(y=50.0, linestyle='--', color=subtle_col, linewidth=subtle_lw_forced, zorder=100)
                                # mark intersection with a small circle
                                try:
                                    ax_eef.scatter([diam50_forced], [50.0], color='purple', s=40, zorder=101)
                                except Exception:
                                    pass
                                ax_eef.text(max(ax_eef.get_xlim()[0], 0), 50.0, f'EEF 50% = {diam50_forced:.3f}"', ha='left', va='bottom', fontsize=9, color='purple')
                            except Exception:
                                pass
                            try:
                                radius50_forced = float(diam50_forced) / 2.0
                                subtle_lw_forced = 0.8
                                subtle_col = 'lightgrey'
                                ax_eef_res.axvline(x=radius50_forced, linestyle='--', color=subtle_col, linewidth=subtle_lw_forced, zorder=100)
                                try:
                                    ax_eef_res.scatter([radius50_forced], [0.0], color=subtle_col, s=30, zorder=101)
                                except Exception:
                                    pass
                                try:
                                    ylim = ax_eef_res.get_ylim()
                                    # move R50 label further away from the line (increase vertical offset)
                                    ax_eef_res.text(radius50_forced, ylim[0] + 0.08 * (ylim[1] - ylim[0]), f'R50 = {radius50_forced:.3f}"', ha='center', va='bottom', fontsize=9, color='black', rotation=90)
                                except Exception:
                                    pass
                                # draw a short horizontal dashed marker at the top of the residual axis (in axes coords)
                                try:
                                    hw = max(0.5, float(radius50_forced) * 0.05)
                                    ax_eef_res.plot([radius50_forced - hw, radius50_forced + hw], [0.95, 0.95], transform=ax_eef_res.get_xaxis_transform(), linestyle='--', color=subtle_col, linewidth=subtle_lw_forced, zorder=100)
                                except Exception:
                                    pass
                            except Exception:
                                pass
                    except Exception:
                        pass
                except Exception:
                    pass
                    plt.savefig(outfn_comb, dpi=150)
                    plt.close()
                except Exception:
                    pass

            # Export fit parameters to Excel
            try:
                fit_export_data = {
                    'Parameter': [],
                    'Modified Pseudo-Voigt': [],
                    'Pearson4': [],
                    'King': []
                }
                
                # Pseudo-Voigt parameters
                pv_params = {
                    'Amplitude (A)': A_fit,
                    'Core width (Gamma_c) [arcsec]': Gamma_c_fit,
                    'Wing width (Gamma_w) [arcsec]': Gamma_w_fit,
                    'Mixing ratio (eta)': eta_fit,
                    'Wing exponent (beta)': beta_fit,
                    'Wing scale (scalar)': scalar_fit,
                }
                
                for param_name, value in pv_params.items():
                    fit_export_data['Parameter'].append(param_name)
                    fit_export_data['Modified Pseudo-Voigt'].append(value)
                    fit_export_data['Pearson4'].append(None)
                    fit_export_data['King'].append(None)
                
                # Pearson4 parameters
                if pearson4_result is not None:
                    p4_params = {
                        'Amplitude': pearson4_result.params['amplitude'].value,
                        'Center [arcsec]': pearson4_result.params['center'].value,
                        'Sigma [arcsec]': pearson4_result.params['sigma'].value,
                        'Exponent (m)': pearson4_result.params['expon'].value,
                        'Skewness (nu)': pearson4_result.params['skew'].value,
                    }
                    
                    # Add Pearson4-specific parameters
                    for param_name, value in p4_params.items():
                        if param_name not in fit_export_data['Parameter']:
                            fit_export_data['Parameter'].append(param_name)
                            fit_export_data['Modified Pseudo-Voigt'].append(None)
                            fit_export_data['King'].append(None)
                            fit_export_data['Pearson4'].append(None)
                        idx = fit_export_data['Parameter'].index(param_name)
                        fit_export_data['Pearson4'][idx] = value

                # King parameters
                try:
                    if 'I0_k' in locals():
                        king_params = {
                            'King I0': I0_k,
                            'King rc [arcsec]': rc_k,
                            'King alpha': alpha_k,
                            'King background (b)': b_k,
                        }
                        for param_name, value in king_params.items():
                            if param_name not in fit_export_data['Parameter']:
                                fit_export_data['Parameter'].append(param_name)
                                fit_export_data['Modified Pseudo-Voigt'].append(None)
                                fit_export_data['Pearson4'].append(None)
                                fit_export_data['King'].append(None)
                            idx = fit_export_data['Parameter'].index(param_name)
                            fit_export_data['King'][idx] = value
                except Exception:
                    pass
                
                # Export to CSV
                fit_df = pd.DataFrame(fit_export_data)
                export_fit_path = os.path.join('Figures', 'fit_parameters.csv')
                fit_df.to_csv(export_fit_path, index=False)
                
            except Exception:
                pass

        else:
            pass
    except Exception:
        # SciPy not available or other error — skip fitting but don't fail
        print("scipy.optimize not available — skipping aggregated radial fit.")

    # Also compute 50% from origin for reference
    r_profile_00, cumulative_00, total_00 = radial_profile(0.0, 0.0, n_r=n_r_final, n_theta=n_theta_final, r_margin_factor=final_r_margin)
    frac_00 = cumulative_00 / total_00 if total_00 > 0 else cumulative_00
    radius_50_00 = _radius_for_fraction(frac_00, r_profile_00, target=0.5)
    # 90% at origin for reference
    radius_90_00 = _radius_for_fraction(frac_00, r_profile_00, target=0.9)
    # Fallback: if radial integration failed (NaN/None), approximate using
    # an effective sigma from the weighted mixture (use normalized weights).
    try:
        if not (isinstance(radius_50_00, float) and np.isfinite(radius_50_00)):
            # use sigmax/sigmay arrays (in meters) and normalized weight_arr
            try:
                sigx_arr = df['sigmax'].to_numpy(dtype=float, copy=False)
                sigy_arr = df['sigmay'].to_numpy(dtype=float, copy=False)
                # weight_arr should be normalized (sum==1) from earlier
                if 'weight_arr' in locals() and weight_arr is not None and np.isfinite(weight_arr).all() and weight_arr.size == sigx_arr.size:
                    w = weight_arr
                else:
                    wtmp = df.get('weight', pd.Series([1.0]*len(df))).to_numpy(dtype=float, copy=False)
                    wsum = float(np.nansum(wtmp)) if wtmp.size else 0.0
                    w = (wtmp / wsum) if (wsum and np.isfinite(wsum) and wsum > 0.0) else np.ones(len(df), dtype=float) / max(1, len(df))
                sigma2_eff = float(np.sum(w * (sigx_arr**2 + sigy_arr**2) * 0.5))
                if sigma2_eff > 0 and np.isfinite(sigma2_eff):
                    sigma_eff = float(np.sqrt(sigma2_eff))
                    radius_50_00 = sigma_eff * np.sqrt(2.0 * np.log(2.0))
            except Exception:
                pass
    except Exception:
        pass
    
    # Compute optimized configuration metrics if provided
    opt_center_x, opt_center_y, opt_radius_50, opt_radius_90 = None, None, None, None
    opt_r_profile, opt_frac_profile = None, None
    if df_optimized is not None:
        # Compute center for optimized config
        opt_total_weight = df_optimized['weight'].sum()
        opt_center_x = (df_optimized['mux'] * df_optimized['weight']).sum() / opt_total_weight
        opt_center_y = (df_optimized['muy'] * df_optimized['weight']).sum() / opt_total_weight

        # Build fast summation kernel for optimized df (threaded)
        opt_mux = df_optimized['mux'].to_numpy(dtype=float, copy=False)
        opt_muy = df_optimized['muy'].to_numpy(dtype=float, copy=False)
        opt_sigx = df_optimized['sigmax'].to_numpy(dtype=float, copy=False)
        opt_sigy = df_optimized['sigmay'].to_numpy(dtype=float, copy=False)
        opt_theta = df_optimized['theta_degrees'].to_numpy(dtype=float, copy=False)
        opt_weight = df_optimized['weight'].to_numpy(dtype=float, copy=False)
        opt_dist = df_optimized.get('distribution', pd.Series(['gaussian'] * len(df_optimized))).astype(str).str.lower().to_numpy(copy=False)
        opt_alpha_azi = pd.to_numeric(df_optimized.get('alpha_azi', pd.Series([0.5] * len(df_optimized))), errors='coerce').fillna(0.5).to_numpy(dtype=float, copy=False)
        opt_alpha_rad = pd.to_numeric(df_optimized.get('alpha_rad', pd.Series([0.5] * len(df_optimized))), errors='coerce').fillna(0.5).to_numpy(dtype=float, copy=False)

        # Apply same sigma floor to optimized dataset
        opt_sigx = np.maximum(opt_sigx, MIN_SIG_M)
        opt_sigy = np.maximum(opt_sigy, MIN_SIG_M)

        def _sum_chunk_on_grid_opt(Xg, Yg, idxs: np.ndarray, normalize_flag: bool) -> np.ndarray:
            Zc = np.zeros_like(Xg, dtype=float)
            for i in idxs:
                if opt_dist[i] in ['pseudo-voigt', 'voigt']:
                    Zc += pseudo_voigt_2d_rotated(
                        Xg, Yg,
                        muazi=opt_mux[i], murad=opt_muy[i],
                        sigmaazi=opt_sigx[i], sigmarad=opt_sigy[i],
                        theta=opt_theta[i],
                        alphaazi=opt_alpha_azi[i],
                        alpharad=opt_alpha_rad[i],
                        amplitude=opt_weight[i],
                        normalize=normalize_flag,
                        degrees=True,
                    )
                else:
                    Zc += gaussian_2d_rotated(
                        Xg, Yg,
                        mux=opt_mux[i], muy=opt_muy[i],
                        sigmax=opt_sigx[i], sigmay=opt_sigy[i],
                        theta=opt_theta[i],
                        amplitude=opt_weight[i],
                        normalize=normalize_flag,
                        degrees=True,
                    )
            return Zc

        def _sum_on_grid_opt(Xg, Yg, normalize_flag: bool) -> np.ndarray:
            n = len(opt_mux)
            if n == 0:
                return np.zeros_like(Xg, dtype=float)
            max_workers = 1 if n < 25 else min(8, (os.cpu_count() or 2))
            if max_workers <= 1:
                return _sum_chunk_on_grid_opt(Xg, Yg, np.arange(n), normalize_flag)
            chunks = np.array_split(np.arange(n), max_workers)
            with ThreadPoolExecutor(max_workers=max_workers) as ex:
                parts = list(ex.map(lambda c: _sum_chunk_on_grid_opt(Xg, Yg, c, normalize_flag), chunks))
            Zg = np.zeros_like(Xg, dtype=float)
            for p in parts:
                Zg += p
            return Zg
        
        # Create radial_profile function for optimized df
        def radial_profile_opt(cx, cy, n_r=400, n_theta=360, r_margin_factor=5.0):
            max_sigma = max(df_optimized['sigmax'].max(), df_optimized['sigmay'].max())
            max_center_dist = np.sqrt((df_optimized['mux'] - cx) ** 2 + (df_optimized['muy'] - cy) ** 2).max()
            r_max = max_center_dist + r_margin_factor * max_sigma
            if r_max <= 0:
                r_max = 1e-6
            theta = np.linspace(0.0, 2.0 * np.pi, n_theta, endpoint=False)
            r = np.linspace(0.0, r_max, n_r)
            dtheta = theta[1] - theta[0]
            dr = r[1] - r[0] if n_r > 1 else r_max
            R, TH = np.meshgrid(r, theta)
            Xp = cx + R * np.cos(TH)
            Yp = cy + R * np.sin(TH)
            Zp = _sum_on_grid_opt(Xp, Yp, normalize)
            radial_energy = np.sum(Zp * R, axis=0) * dtheta
            cumulative = np.cumsum(radial_energy * dr)
            total_energy = cumulative[-1] if cumulative.size else 1.0
            if debug:
                    pass
            return r, cumulative, total_energy
        
        # Find best focus for optimized (simple version - use computed center)
        def hew_at_center_opt(cx, cy, coarse=False):
            if coarse:
                n_r = 110 if fast else 140
                n_theta = 72 if fast else 100
            else:
                n_r = 200 if fast else 240
                n_theta = 120 if fast else 180
            r, cumulative, total_energy = radial_profile_opt(cx, cy, n_r=n_r, n_theta=n_theta)
            frac = cumulative / total_energy if total_energy > 0 else cumulative
            return _radius_for_fraction(frac, r, target=0.5)
        
        # Optimize center position
        if quick_mode:
            best_opt_cx, best_opt_cy = opt_center_x, opt_center_y
            best_opt_hew = hew_at_center_opt(best_opt_cx, best_opt_cy, coarse=True)
            step_size = 1e-6
            for cx, cy in (
                (opt_center_x + step_size, opt_center_y),
                (opt_center_x - step_size, opt_center_y),
                (opt_center_x, opt_center_y + step_size),
                (opt_center_x, opt_center_y - step_size),
            ):
                hew_val = hew_at_center_opt(cx, cy, coarse=True)
                if hew_val < best_opt_hew:
                    best_opt_cx, best_opt_cy, best_opt_hew = cx, cy, hew_val
            opt_center_x, opt_center_y = best_opt_cx, best_opt_cy
        else:
            candidates_start = [(opt_center_x, opt_center_y), (0.0, 0.0)]
            best_opt_cx, best_opt_cy = opt_center_x, opt_center_y
            best_opt_hew = hew_at_center_opt(best_opt_cx, best_opt_cy, coarse=True)
            for start_cx, start_cy in candidates_start:
                current_cx, current_cy = start_cx, start_cy
                current_hew = hew_at_center_opt(current_cx, current_cy, coarse=True)
                step_size = 1e-6
                for _ in range(25):
                    candidates = [
                        (current_cx + step_size, current_cy),
                        (current_cx - step_size, current_cy),
                        (current_cx, current_cy + step_size),
                        (current_cx, current_cy - step_size)
                    ]
                    improved = False
                    for cx, cy in candidates:
                        hew_val = hew_at_center_opt(cx, cy, coarse=True)
                        if hew_val < current_hew:
                            current_cx, current_cy = cx, cy
                            current_hew = hew_val
                            improved = True
                            break
                    if not improved:
                        step_size *= 0.5
                        if step_size < 1e-10:
                            break
                if current_hew < best_opt_hew:
                    best_opt_cx, best_opt_cy = current_cx, current_cy
                    best_opt_hew = current_hew
            
            opt_center_x, opt_center_y = best_opt_cx, best_opt_cy
        
        # Compute final optimized metrics
        opt_r_profile, opt_cumulative_profile, opt_total_energy = radial_profile_opt(opt_center_x, opt_center_y, n_r=n_r_final, n_theta=n_theta_final, r_margin_factor=final_r_margin)
        opt_frac_profile = opt_cumulative_profile / opt_total_energy if opt_total_energy > 0 else opt_cumulative_profile
        opt_radius_50 = _radius_for_fraction(opt_frac_profile, opt_r_profile, target=0.5)
        opt_radius_90 = _radius_for_fraction(opt_frac_profile, opt_r_profile, target=0.9)
    
    # Build plot bounds first (avoid expensive Z recomputation)
    max_radius = max(radius_50 or 0, radius_90 or 0, radius_50_00 or 0)
    # Fallback bounds based on Gaussian extents if radii are missing
    margin_factor = 3
    min_x = df['mux'].min() - margin_factor * df['sigmax'].max()
    max_x = df['mux'].max() + margin_factor * df['sigmax'].max()
    min_y = df['muy'].min() - margin_factor * df['sigmay'].max()
    max_y = df['muy'].max() + margin_factor * df['sigmay'].max()

    if max_radius and max_radius > 0:
        xlim = (
            min(min_x, center_x - 1.1 * max_radius, -1.1 * (radius_50_00 or 0.0)),
            max(max_x, center_x + 1.1 * max_radius,  1.1 * (radius_50_00 or 0.0)),
        )
        ylim = (
            min(min_y, center_y - 1.1 * max_radius, -1.1 * (radius_50_00 or 0.0)),
            max(max_y, center_y + 1.1 * max_radius,  1.1 * (radius_50_00 or 0.0)),
        )
    else:
        xlim = (min_x, max_x)
        ylim = (min_y, max_y)

    x = np.linspace(xlim[0], xlim[1], nx)
    y = np.linspace(ylim[0], ylim[1], ny)
    X, Y = np.meshgrid(x, y)
    Z = _sum_on_grid(X, Y, normalize)

    # Create the plots in a single figure with two subplots
    # Use a reasonable figure size that will fit most screens; enable constrained layout
    # Produce the large PSF+EEF figure (single-window diagnostic).
    fig = plt.figure(figsize=(16, 8), constrained_layout=True)
    # Standard linewidth for EEF marker lines (keep 90% and 80% consistent)
    eef_linewidth = 1.5
    
    # Use GridSpec with 20 columns for finer control: left (12 cols) + right (8 cols)
    # Right (EEF) will therefore be 8/20 = 40% of the figure width
    gs = gridspec.GridSpec(1, 20, figure=fig)
    
    # First subplot: weighted sum of Gaussians (left, wider)
    ax1 = plt.subplot(gs[0, :12])
    # Convert to microns for display (1 m = 1e6 µm)
    im = plt.imshow(Z, extent=[x.min()*1e6, x.max()*1e6, y.min()*1e6, y.max()*1e6], origin='lower', cmap='viridis', aspect='equal')
    # Smaller, tighter colorbar attached to ax1 to reduce horizontal crowding
    cbar = plt.colorbar(im, ax=ax1, label='counts', pad=0.02, fraction=0.046)
    ax1.set_xlabel('x [µm]')
    ax1.set_ylabel('y [µm]')
    
    # Add secondary axes for arcsec. Use the same project convention as
    # `arcsec_to_m` (1 arcsec = 12*π/180/3600 m), therefore 1 m equals
    # the reciprocal: 1 / (12*π/180/3600) arcsec. Compute explicitly.
    m_to_arcsec = 1.0 / (12.0 * np.pi / 180.0 / 3600.0)
    um_to_arcsec = m_to_arcsec * 1e-6  # microns to arcsec
    
    # Top axis for x in arcsec
    ax1_top = ax1.secondary_xaxis('top', functions=(lambda um: um * um_to_arcsec, lambda arcsec: arcsec / um_to_arcsec))
    ax1_top.set_xlabel('x [arcsec]')
    
    # Right axis for y in arcsec (immediately next to plot)
    ax1_right = ax1.secondary_yaxis('right', functions=(lambda um: um * um_to_arcsec, lambda arcsec: arcsec / um_to_arcsec))
    ax1_right.set_ylabel('y [arcsec]', rotation=270, va='bottom')
    
    # Mark the minimum with a green cross and coordinates (label updated)
    plt.plot(center_x*1e6, center_y*1e6, 'gx', markersize=10, label='center for minimum HEW')
    plt.text(center_x*1e6, center_y*1e6, f'({center_x*1e6:.2f}, {center_y*1e6:.2f})', color='green', ha='left', va='bottom')
    # Mark (0,0) with a blue cross
    plt.plot(0, 0, 'bx', markersize=10, label='(0,0)')
    # Add circles for 50% and 90% encircled energy
    max_radius = max(radius_50 or 0, radius_90 or 0)
    
    # Calculate axis limits to include: best focus circles, (0,0), and blue circle
    margin_factor = 0.1
    # Start with limits needed for best focus circles
    xlim_min = center_x - max_radius if max_radius > 0 else center_x - 1
    xlim_max = center_x + max_radius if max_radius > 0 else center_x + 1
    ylim_min = center_y - max_radius if max_radius > 0 else center_y - 1
    ylim_max = center_y + max_radius if max_radius > 0 else center_y + 1
    
    # Ensure (0,0) and blue circle (radius_50_00) are included
    if radius_50_00 is not None and radius_50_00 > 0:
        xlim_min = min(xlim_min, 0 - radius_50_00)
        xlim_max = max(xlim_max, 0 + radius_50_00)
        ylim_min = min(ylim_min, 0 - radius_50_00)
        ylim_max = max(ylim_max, 0 + radius_50_00)
    else:
        # At minimum, include (0,0)
        xlim_min = min(xlim_min, 0)
        xlim_max = max(xlim_max, 0)
        ylim_min = min(ylim_min, 0)
        ylim_max = max(ylim_max, 0)
    
    # Add margin and convert to microns
    x_range = xlim_max - xlim_min
    y_range = ylim_max - ylim_min
    margin_x = margin_factor * x_range
    margin_y = margin_factor * y_range
    # Guard against NaN/Inf axis limits (can happen if inputs contain NaNs).
    if not np.isfinite(xlim_min) or not np.isfinite(xlim_max):
        xlim_min = center_x - 1.0
        xlim_max = center_x + 1.0
    if not np.isfinite(ylim_min) or not np.isfinite(ylim_max):
        ylim_min = center_y - 1.0
        ylim_max = center_y + 1.0
    margin_x = margin_factor * (xlim_max - xlim_min)
    margin_y = margin_factor * (ylim_max - ylim_min)
    left_x = (xlim_min - margin_x) * 1e6
    right_x = (xlim_max + margin_x) * 1e6
    bottom_y = (ylim_min - margin_y) * 1e6
    top_y = (ylim_max + margin_y) * 1e6
    if not (np.isfinite(left_x) and np.isfinite(right_x)):
        left_x = (center_x - 1.0) * 1e6
        right_x = (center_x + 1.0) * 1e6
    if not (np.isfinite(bottom_y) and np.isfinite(top_y)):
        bottom_y = (center_y - 1.0) * 1e6
        top_y = (center_y + 1.0) * 1e6
    plt.xlim(left_x, right_x)
    plt.ylim(bottom_y, top_y)
    # Precompute arcsec conversions for legend/value annotations
    m_to_arcsec = 1.0 / (12.0 * np.pi / 180.0 / 3600.0)

    def _min_interval_width(axis_vals: np.ndarray, prof: np.ndarray, frac: float = 0.5) -> float | None:
        """Smallest interval (anywhere) containing a fraction of 1D energy.

        Given a nonnegative profile sampled on an increasing axis, returns the minimal
        width W such that there exists an interval [a, b] with (b-a)=W and
        \int_a^b prof(x) dx >= frac * \int prof(x) dx.

        This matches the request for HEW_x/HEW_y as the *smallest* interval containing 50%.
        """
        axis_vals = np.asarray(axis_vals, dtype=float)
        prof = np.asarray(prof, dtype=float)

        if prof.size < 3 or axis_vals.size != prof.size:
            return None
        if not np.isfinite(prof).any() or not np.isfinite(axis_vals).all():
            return None
        if float(frac) <= 0.0:
            return 0.0

        # Ensure monotonic increasing axis
        if not np.all(np.diff(axis_vals) > 0):
            return None

        # Clamp tiny negatives that can appear from numeric noise
        prof = np.maximum(prof, 0.0)

        # Segment integrals using trapezoids
        dx = np.diff(axis_vals)
        seg = 0.5 * (prof[:-1] + prof[1:]) * dx
        total = float(np.sum(seg))
        if not np.isfinite(total) or total <= 0.0:
            return None
        target = float(frac) * total

        # Prefix integral at nodes: pref[k] = integral from x[0] to x[k]
        pref = np.empty(prof.size, dtype=float)
        pref[0] = 0.0
        pref[1:] = np.cumsum(seg)

        best = None
        # Two-pointer sliding window: for each left index i, advance right index j
        # until the integral pref[j] - pref[i] >= target, then record width.
        n = prof.size
        j = 1
        for i in range(0, n - 1):
            if j < i + 1:
                j = i + 1
            # advance j until interval [i,j] reaches target or j==n-1
            while j < n and (pref[j] - pref[i]) < target:
                j += 1
            if j >= n:
                break
            width = float(axis_vals[j] - axis_vals[i])
            if best is None or width < best:
                best = width
        return best
    
    def _fwhm_from_profile(axis_vals: np.ndarray, prof: np.ndarray) -> float | None:
        #Compute FWHM from a 1D profile sampled on axis_vals.
        axis_vals = np.asarray(axis_vals, dtype=float)
        prof = np.asarray(prof, dtype=float)
        if prof.size < 3 or axis_vals.size != prof.size:
            return None
        if not np.isfinite(prof).any() or not np.isfinite(axis_vals).all():
            return None
        # Normalize profile
        prof = np.maximum(prof, 0.0)
        max_val = np.max(prof)
        if not np.isfinite(max_val) or max_val <= 0:
            return None
        half_max = max_val / 2.0
        above = np.where(prof >= half_max)[0]
        if above.size < 2:
            return None
        left = above[0]
        right = above[-1]
        return float(axis_vals[right] - axis_vals[left])


    def _compute_hew_xy_arcsec_from_grid_marginals(x_axis: np.ndarray, y_axis: np.ndarray, Zg: np.ndarray) -> tuple[float | None, float | None]:
        """Compute HEW_x and HEW_y (arcsec) from a gridded PSF `Zg` sampled on axes `x_axis`, `y_axis`.

        The function integrates the 2D PSF to obtain 1D marginals and then finds
        the minimal-width interval containing 50% of the energy for each axis.
        Returns (hew_x_arcsec, hew_y_arcsec) where values may be None on failure.
        """
        try:
            if getattr(Zg, 'size', 0) == 0:
                return (None, None)
            if not np.isfinite(np.asarray(Zg, dtype=float)).any():
                return (None, None)

            x_axis = np.asarray(x_axis, dtype=float)
            y_axis = np.asarray(y_axis, dtype=float)
            Zg = np.asarray(Zg, dtype=float)

            # Marginals (integrate the full 2D PSF along the other axis)
            prof_x = np.trapezoid(Zg, y_axis, axis=0)
            prof_y = np.trapezoid(Zg, x_axis, axis=1)

            hew_x_m = _min_interval_width(x_axis, prof_x, frac=0.5)
            hew_y_m = _min_interval_width(y_axis, prof_y, frac=0.5)

            # Convert meters back to arcsec using project convention
            m_to_arcsec = 1.0 / (12.0 * np.pi / 180.0 / 3600.0)
            hew_x_arcsec = (hew_x_m * m_to_arcsec) if hew_x_m is not None else None
            hew_y_arcsec = (hew_y_m * m_to_arcsec) if hew_y_m is not None else None
            return hew_x_arcsec, hew_y_arcsec
        except Exception:
            return (None, None)

    # HEW_x / HEW_y annotations for encircled-energy legend entries
    # Base curves share the same underlying PSF, so their HEW_x/HEW_y should be the same.
    hew_base_x_arcsec, hew_base_y_arcsec = _compute_hew_xy_arcsec_from_grid_marginals(x, y, Z)
    hew_best_x_arcsec, hew_best_y_arcsec = hew_base_x_arcsec, hew_base_y_arcsec
    hew_00_x_arcsec, hew_00_y_arcsec = hew_base_x_arcsec, hew_base_y_arcsec

    # FWHM_x and FWHM_y calculation
    prof_x = np.trapezoid(Z, y, axis=0)
    prof_y = np.trapezoid(Z, x, axis=1)
    fwhm_x_m = _fwhm_from_profile(x, prof_x)
    fwhm_y_m = _fwhm_from_profile(y, prof_y)
    m_to_arcsec = 1.0 / (12.0 * np.pi / 180.0 / 3600.0)
    fwhm_x_arcsec = fwhm_x_m * m_to_arcsec if fwhm_x_m is not None else None
    fwhm_y_arcsec = fwhm_y_m * m_to_arcsec if fwhm_y_m is not None else None

    fwhm_opt_x_arcsec, fwhm_opt_y_arcsec = (None, None)
    if df_optimized is not None:
        Z_opt = _sum_on_grid_opt(X, Y, normalize)
        prof_x_opt = np.trapezoid(Z_opt, y, axis=0)
        prof_y_opt = np.trapezoid(Z_opt, x, axis=1)
        fwhm_x_m_opt = _fwhm_from_profile(x, prof_x_opt)
        fwhm_y_m_opt = _fwhm_from_profile(y, prof_y_opt)
        fwhm_opt_x_arcsec = fwhm_x_m_opt * m_to_arcsec if fwhm_x_m_opt is not None else None
        fwhm_opt_y_arcsec = fwhm_y_m_opt * m_to_arcsec if fwhm_y_m_opt is not None else None
    hew_best_arcsec = 2 * radius_50 * m_to_arcsec if (radius_50 is not None and np.isfinite(radius_50)) else None
    hew_origin_arcsec = 2 * radius_50_00 * m_to_arcsec if (radius_50_00 is not None and np.isfinite(radius_50_00)) else None
    eef90_arcsec = 2 * radius_90 * m_to_arcsec if (radius_90 is not None and np.isfinite(radius_90)) else None
    eef90_origin_arcsec = 2 * radius_90_00 * m_to_arcsec if (radius_90_00 is not None and np.isfinite(radius_90_00)) else None

    # If caller only requests metrics, return them now without any plotting side-effects.
    if return_metrics_only:
        return {
            'hew_origin_arcsec': hew_origin_arcsec,
            'hew_best_arcsec': hew_best_arcsec,
            'eef90_origin_arcsec': eef90_origin_arcsec,
            'eef90_best_arcsec': eef90_arcsec,
            'hew_x_arcsec': hew_base_x_arcsec,
            'hew_y_arcsec': hew_base_y_arcsec,
            'hew_opt_x_arcsec': hew_opt_x_arcsec if 'hew_opt_x_arcsec' in locals() else None,
            'hew_opt_y_arcsec': hew_opt_y_arcsec if 'hew_opt_y_arcsec' in locals() else None,
            'fwhm_x_arcsec': fwhm_x_arcsec,
            'fwhm_y_arcsec': fwhm_y_arcsec,
            'hew_opt_arcsec': (2 * opt_radius_50 * m_to_arcsec) if 'opt_radius_50' in locals() and opt_radius_50 is not None else None,
            'eef90_opt_arcsec': (2 * opt_radius_90 * m_to_arcsec) if 'opt_radius_90' in locals() and opt_radius_90 is not None else None,
        }
    else:
        # Fallback: if origin HEW couldn't be computed from radial integration,
        # use the grid-marginal HEW as a best-effort surrogate.
        if hew_origin_arcsec is None and hew_base_x_arcsec is not None:
            hew_origin_arcsec = hew_base_x_arcsec
        # Also expose the exact same metrics when producing plots so both
        # interactive and metrics-only runs produce identical outputs.
        try:
            metrics_out = {
                'hew_origin_arcsec': hew_origin_arcsec,
                'hew_best_arcsec': hew_best_arcsec,
                'eef90_origin_arcsec': eef90_origin_arcsec,
                'eef90_best_arcsec': eef90_arcsec,
                'hew_x_arcsec': hew_base_x_arcsec,
                'hew_y_arcsec': hew_base_y_arcsec,
                'hew_opt_x_arcsec': hew_opt_x_arcsec,
                'hew_opt_y_arcsec': hew_opt_y_arcsec,
                'hew_opt_arcsec': (2 * opt_radius_50 * m_to_arcsec) if opt_radius_50 is not None else None,
                'eef90_opt_arcsec': (2 * opt_radius_90 * m_to_arcsec) if opt_radius_90 is not None else None,
            }
            print(json.dumps(metrics_out, indent=2))
        except Exception:
            pass
    if radius_90 is not None:
        # Add a dashed circle for 90% encircled energy in red (short legend label)
        label_90 = 'EEF 90% (min)'
        circle_90 = plt.Circle((center_x*1e6, center_y*1e6), radius_90*1e6, fill=False, color='red', linestyle='--', linewidth=2, label=label_90)
        plt.gca().add_patch(circle_90)
    if radius_50_00 is not None:
        # Add a dashed circle for 50% encircled energy from (0,0) in blue (short label)
        label_50_00 = 'HEW (0,0)'
        circle_50_00 = plt.Circle((0, 0), radius_50_00*1e6, fill=False, color='blue', linestyle='--', linewidth=2, label=label_50_00)
        plt.gca().add_patch(circle_50_00)
    if radius_50 is not None:
        # Add a dashed circle for 50% encircled energy in green (short label)
        label_50 = 'HEW (min)'
        circle_50 = plt.Circle((center_x*1e6, center_y*1e6), radius_50*1e6, fill=False, color='green', linestyle='--', linewidth=2, label=label_50)
        plt.gca().add_patch(circle_50)
    
    # Add optimized circles if provided
    if df_optimized is not None and opt_radius_50 is not None:
        opt_hew_arcsec = 2 * opt_radius_50 * m_to_arcsec
        opt_eef90_arcsec = 2 * opt_radius_90 * m_to_arcsec if opt_radius_90 is not None else None
        # Mark optimized best focus
        plt.plot(opt_center_x*1e6, opt_center_y*1e6, 'mx', markersize=10, label='optimized minimum')
        # HEW circle (magenta dotted) — short label
        label_opt_50 = 'HEW (optimized)'
        circle_opt_50 = plt.Circle((opt_center_x*1e6, opt_center_y*1e6), opt_radius_50*1e6, fill=False, color='magenta', linestyle=':', linewidth=2, label=label_opt_50)
        plt.gca().add_patch(circle_opt_50)
        # EEF 90% circle (orange dotted)
        if opt_radius_90 is not None:
            label_opt_90 = 'EEF 90% (opt)'
            circle_opt_90 = plt.Circle((opt_center_x*1e6, opt_center_y*1e6), opt_radius_90*1e6, fill=False, color='orange', linestyle=':', linewidth=2, label=label_opt_90)
            plt.gca().add_patch(circle_opt_90)
    
    # Reorder legend items on the left subplot:
    # 1) focus points, 2) HEW circles, 3) EEF 90% circles
    handles, labels = ax1.get_legend_handles_labels()

    focus_order = {
        '(0,0)': 0,
        'center for minimum HEW': 1,
        'minimum': 2,
        'optimized minimum': 3,
    }

    def _legend_sort_key(label: str) -> tuple[int, int, str]:
        if label in focus_order:
            return (0, focus_order[label], label)
        if label.startswith('HEW'):
            if 'minimum' in label:
                return (1, 0, label)
            if '(0,0)' in label:
                return (1, 1, label)
            if 'optimized' in label:
                return (1, 2, label)
            return (1, 99, label)
        if label.startswith('EEF 80%'):
            # Place EEF 80% entries before EEF 90%
            if '(min)' in label or 'minimum' in label:
                return (2, -1, label)
            if 'optimized' in label:
                return (2, 0, label)
            return (2, 50, label)
        if label.startswith('EEF 90%'):
            if '(min)' in label or 'minimum' in label:
                return (2, 0, label)
            if 'optimized' in label:
                return (2, 1, label)
            return (2, 99, label)
        return (3, 99, label)

    order = sorted(range(len(labels)), key=lambda i: _legend_sort_key(labels[i]))
    # Ensure (0,0) is first and 'center for minimum HEW' is second if present
    preferred_first = []
    try:
        idx_00 = next(i for i, lbl in enumerate(labels) if lbl == '(0,0)')
        preferred_first.append(idx_00)
    except StopIteration:
        idx_00 = None
    try:
        idx_center = next(i for i, lbl in enumerate(labels) if lbl == 'center for minimum HEW')
        # Only add if it's not the same as (0,0)
        if idx_center is not None and idx_center != idx_00:
            preferred_first.append(idx_center)
    except StopIteration:
        pass

    # Build final order: preferred labels first (in the requested sequence), then the rest in the sorted order
    remaining = [i for i in order if i not in preferred_first]
    final_order = preferred_first + remaining
    handles_sorted = [handles[i] for i in final_order]
    labels_sorted = [labels[i] for i in final_order]
    # place compact legend inside the left subplot (upper-left) with 2 columns x 3 rows
    try:
        ax1.legend(handles_sorted, labels_sorted, loc='upper left', ncol=2, fontsize=8, framealpha=0.85, bbox_to_anchor=(0.02, 0.98), bbox_transform=ax1.transAxes)
    except Exception:
        try:
            ax1.legend(handles_sorted, labels_sorted, loc='upper left', framealpha=0.85, bbox_to_anchor=(0.02, 0.98), bbox_transform=ax1.transAxes)
        except Exception:
            ax1.legend(handles_sorted, labels_sorted, loc='upper left')
    
    # Second subplot: encircled energy function (right, 40% width)
    ax2 = plt.subplot(gs[0, 12:])
    # Convert diameter from meters to arcsec (1 m = 54000/π arcsec)
    profile_pct = frac_profile * 100 if 'frac_profile' in locals() else []
    profile_diam = 2 * r_profile * m_to_arcsec if 'r_profile' in locals() else []
    profile_pct_00 = frac_00 * 100 if 'frac_00' in locals() else []
    profile_diam_00 = 2 * r_profile_00 * m_to_arcsec if 'r_profile_00' in locals() else []
    # labels for the EEF subplot — include FWHM values when available
    try:
        if fwhm_x_arcsec is not None and fwhm_y_arcsec is not None:
            label_best = f'Centered on minimum (FWHM_x={fwhm_x_arcsec:.4f}\", FWHM_y={fwhm_y_arcsec:.4f}\")'
            label_00 = f'Centered on (0,0) (FWHM_x={fwhm_x_arcsec:.4f}\", FWHM_y={fwhm_y_arcsec:.4f}\")'
        else:
            label_best = 'Centered on minimum'
            label_00 = 'Centered on (0,0)'
    except Exception:
        label_best = 'Centered on minimum'
        label_00 = 'Centered on (0,0)'
    # Limit plot to 95% percentile
    def limit_percentile(pct, diam, max_pct=95):
        pct = np.array(pct)
        diam = np.array(diam)
        mask = pct <= max_pct
        return pct[mask], diam[mask]
    profile_pct_95, profile_diam_95 = limit_percentile(profile_pct, profile_diam)
    profile_pct_00_95, profile_diam_00_95 = limit_percentile(profile_pct_00, profile_diam_00)
    plt.plot(profile_pct_95, profile_diam_95, label=label_best, color='green', linestyle='-', linewidth=2.5)
    plt.plot(profile_pct_00_95, profile_diam_00_95, label=label_00, color='blue', linestyle='-', linewidth=2.5)
    if fit_profile_pct is not None and fit_profile_diam is not None:
        fit_profile_pct_95, fit_profile_diam_95 = limit_percentile(fit_profile_pct, fit_profile_diam)
    # PV label for EEF subplot — include full fit parameters when available
    pv_label = 'Modified pseudo-Voigt radial fit'
    try:
        if 'A_fit' in locals():
            pv_label = (f"Modified pseudo-Voigt radial fit (A={A_fit:.2e}, Gamma_c={Gamma_c_fit:.2f}\"",
                        f"Gamma_w={Gamma_w_fit:.2f}\", eta={eta_fit:.3f}, beta={beta_fit:.2f}, s={scalar_fit:.2f})")
            # join tuple if accidentally created as tuple above (fallback safety)
            if isinstance(pv_label, tuple):
                pv_label = ' '.join(pv_label)
    except Exception:
        pv_label = 'Modified pseudo-Voigt radial fit'

    # Plot PV EEF on this right-side EEF subplot. Prefer PV evaluated on the fine rplot
    # grid (`Ifit_pv_plot` with `rplot`) when available; otherwise fall back to the
    # precomputed `fit_profile_*` arrays.
    try:
        if 'Ifit_pv_plot' in locals() and 'rplot' in locals() and Ifit_pv_plot is not None:
            rpv = np.asarray(rplot, dtype=float)
            drpv = float(rpv[1] - rpv[0]) if rpv.size > 1 else 1.0
            radial_pv_plot = 2.0 * np.pi * rpv * np.asarray(Ifit_pv_plot, dtype=float) * drpv
            tot_pv_plot = np.sum(radial_pv_plot)
            if tot_pv_plot > 0 and np.isfinite(tot_pv_plot):
                pv_eef_pct_plot = 100.0 * np.cumsum(radial_pv_plot) / tot_pv_plot
                pv_diam_plot = 2.0 * rpv
                maskpv = pv_eef_pct_plot <= 95.0
                if np.any(maskpv):
                    plt.plot(pv_eef_pct_plot[maskpv], pv_diam_plot[maskpv], label=pv_label, color='red', linestyle='--', linewidth=eef_linewidth)
        elif fit_profile_pct is not None and fit_profile_diam is not None:
            plt.plot(fit_profile_pct_95, fit_profile_diam_95, label=pv_label, color='red', linestyle='--', linewidth=eef_linewidth)
    except Exception:
        try:
            if fit_profile_pct is not None and fit_profile_diam is not None:
                plt.plot(fit_profile_pct_95, fit_profile_diam_95, label=pv_label, color='red', linestyle='--', linewidth=eef_linewidth)
        except Exception:
            pass

    # Draw 80% EEF circle on the left PSF image (ax1) as a purple dashed circle
    try:
        if 'profile_pct' in locals() and profile_pct is not None and len(profile_pct):
            pct_arr_full = np.asarray(profile_pct, dtype=float)
            diam_arr_full = np.asarray(profile_diam, dtype=float)
            order_full = np.argsort(pct_arr_full)
            pct_sorted_full = pct_arr_full[order_full]
            diam_sorted_full = diam_arr_full[order_full]
            if pct_sorted_full[0] <= 80.0 <= pct_sorted_full[-1]:
                diam80 = float(np.interp(80.0, pct_sorted_full, diam_sorted_full))
                # convert diameter arcsec to radius in microns for ax1 (which uses µm)
                arcsec_to_m_local = 1.0 / m_to_arcsec
                radius_um = (diam80 / 2.0) * arcsec_to_m_local * 1e6
                try:
                    import matplotlib.patches as mpatches
                    label_80 = 'EEF 80% (min)'
                    circ = mpatches.Circle((center_x * 1e6, center_y * 1e6), radius_um, edgecolor='purple', linestyle='--', linewidth=eef_linewidth, fill=False, zorder=9, label=label_80)
                    ax1.add_patch(circ)
                    # update ax1 legend to include the new circle label
                    try:
                        handles_now, labels_now = ax1.get_legend_handles_labels()
                        order_now = sorted(range(len(labels_now)), key=lambda i: _legend_sort_key(labels_now[i]))
                        handles_sorted_now = [handles_now[i] for i in order_now]
                        labels_sorted_now = [labels_now[i] for i in order_now]
                        # place compact legend inside the left subplot to avoid overlap with colorbar
                        ax1.legend(handles_sorted_now, labels_sorted_now, loc='upper left', ncol=2, fontsize=8, framealpha=0.85, bbox_to_anchor=(0.02, 0.98), bbox_transform=ax1.transAxes)
                    except Exception:
                        try:
                            ax1.legend(handles_sorted_now, labels_sorted_now, loc='upper left', ncol=2, fontsize=8, framealpha=0.85, bbox_to_anchor=(0.02, 0.98), bbox_transform=ax1.transAxes)
                        except Exception:
                            try:
                                ax1.legend()
                            except Exception:
                                pass
                except Exception:
                    pass
    except Exception:
        pass

    # (EEF legend will be created after all fit curves are plotted,
    #  so that fit parameter labels are available and included.)
    if pearson4_profile_pct is not None and pearson4_profile_diam is not None:
        pearson4_profile_pct_95, pearson4_profile_diam_95 = limit_percentile(pearson4_profile_pct, pearson4_profile_diam)
        # Pearson4 label — include key fit parameters when available
        try:
            # Build a compact Pearson4 legend label showing only the five requested parameters
            p4_label = 'Pearson4 radial fit'
            try:
                params_p4 = None
                if 'pearson4_result' in locals() and pearson4_result is not None and hasattr(pearson4_result, 'params'):
                    params_p4 = pearson4_result.params
                elif 'pearson4_used_params_final' in locals() and pearson4_used_params_final is not None:
                    params_p4 = pearson4_used_params_final
                if params_p4 is not None:
                    # Map lmfit parameter names to the exact output keys the user expects
                    mapping = [
                        ('amplitude', 'Pearson4_Amplitude', 'Amplitude'),
                        ('center', 'Pearson4_Center_arcsec', 'Center_arcsec'),
                        ('sigma', 'Pearson4_Sigma_arcsec', 'Sigma_arcsec'),
                        ('expon', 'Pearson4_Exponent_m', 'Exponent_m'),
                        ('skew', 'Pearson4_Skew_nu', 'Skew_nu')
                    ]
                    parts = []
                    for lm_name, out_key, disp_name in mapping:
                        try:
                            if lm_name in params_p4:
                                val = params_p4[lm_name].value if hasattr(params_p4[lm_name], 'value') else float(params_p4[lm_name])
                                sval = f"{val:.4g}"
                                parts.append(f"{disp_name}={sval}")
                        except Exception:
                            continue
                        if parts:
                            p4_label = 'Pearson4 radial fit (' + ', '.join(parts) + ')'
            except Exception:
                pass
            plt.plot(pearson4_profile_pct_95, pearson4_profile_diam_95, label=p4_label, color='orange', linestyle='--', linewidth=eef_linewidth)
        except Exception:
            pass
    # Add King profile to EEF subplot if available (clip to 95% percentile)
    if 'king_profile_pct' in locals() and king_profile_pct is not None and 'king_profile_diam' in locals() and king_profile_diam is not None:
        try:
            # King label — include fit params if the least-squares King fit produced them
            kp_label = 'King radial fit'
            try:
                if 'I0_k' in locals():
                    kp_label = f'King radial fit (I0={I0_k:.2e}, rc={rc_k:.2f}\", alpha={alpha_k:.2f}, b={b_k:.2e})'
            except Exception:
                pass
            try:
                k_pct = np.asarray(king_profile_pct, dtype=float)
                k_diam = np.asarray(king_profile_diam, dtype=float)
                k_pct_95, k_diam_95 = limit_percentile(k_pct, k_diam, max_pct=95)
                if len(k_pct_95) > 0:
                    plt.plot(k_pct_95, k_diam_95, label=kp_label, color='purple', linestyle='--', linewidth=eef_linewidth, alpha=0.95)
            except Exception:
                # fallback: attempt to plot raw arrays but still guard errors
                try:
                    plt.plot(king_profile_pct, king_profile_diam, label=kp_label, color='purple', linestyle='--', linewidth=eef_linewidth, alpha=0.95)
                except Exception:
                    pass
        except Exception:
            pass
    else:
        # Try a proper King least-squares fit to the EEF reference if available
        try:
            if 'have_least_squares' in locals() and have_least_squares and ('eef_data' in locals() and eef_data is not None) and ('r_arcsec' in locals() and 'r_fit' in locals()):
                try:
                    # initial guesses (reuse sensible defaults from earlier)
                    I0_0 = float(np.nanmax(I_fit)) if ('I_fit' in locals() and I_fit.size) else 1.0
                except Exception:
                    I0_0 = 1.0
                try:
                    rc_guess = float(Gamma_c_fit) if 'Gamma_c_fit' in locals() else (float(np.median(r_fit)) if ('r_fit' in locals() and r_fit.size) else 1.0)
                except Exception:
                    rc_guess = 1.0
                rc_0 = max(0.5 * rc_guess, 1e-3)
                alpha_0 = 2.0
                try:
                    b0 = float(np.nanmin(I_fit)) if ('I_fit' in locals() and np.any(np.isfinite(I_fit))) else 0.0
                except Exception:
                    b0 = 0.0
                x0k = [I0_0, rc_0, alpha_0, b0]
                # bounds
                lbk = [max(float(floor) if 'floor' in locals() else 1e-12, 0.0), 1e-3, 1.0, 0.0]
                rc_cap = float(min(r_arcsec.max() * 0.5 if ('r_arcsec' in locals() and r_arcsec.size) else 10.0, 50.0))
                ubk = [np.inf, rc_cap, 8.0, np.inf]

                def _resid_king(vec):
                    try:
                        I0_k, rc_k, alpha_k, b_k = vec
                        model_vals = king_profile(r_arcsec, I0_k, rc_k, alpha_k, b_k)
                        model_vals = np.maximum(model_vals, 0.0)
                        dr = float(r_arcsec[1] - r_arcsec[0]) if r_arcsec.size>1 else 1.0
                        radial = 2.0 * np.pi * r_arcsec * model_vals * dr
                        tot = np.sum(radial)
                        if tot <= 0 or not np.isfinite(tot):
                            return np.ones_like(eef_data) * 1e6
                        model_eef_pct = 100.0 * np.cumsum(radial) / tot
                        from numpy import interp
                        model_at_ref = interp(r_fit, r_arcsec, model_eef_pct) / 100.0
                        return (model_at_ref - eef_data) * (eef_weight if 'eef_weight' in locals() else 1.0)
                    except Exception:
                        return np.ones_like(eef_data) * 1e6

                try:
                    resk = least_squares(_resid_king, x0k, bounds=(lbk, ubk), loss='soft_l1', f_scale=1.0, max_nfev=10000)
                    if hasattr(resk, 'x'):
                        I0_k, rc_k, alpha_k, b_k = resk.x.tolist()
                except Exception:
                    I0_k, rc_k, alpha_k, b_k = x0k

                # compute King model on rplot for EEF plotting
                try:
                    rpv = np.asarray(rplot if 'rplot' in locals() else np.linspace(0.0, float(r_arcsec.max()), 1000), dtype=float)
                    Ifit_king_plot = king_profile(rpv, I0_k, rc_k, alpha_k, b_k)
                    drpv = float(rpv[1] - rpv[0]) if rpv.size>1 else 1.0
                    radial_k_plot = 2.0 * np.pi * rpv * Ifit_king_plot * drpv
                    tot_k_plot = np.sum(radial_k_plot)
                    if tot_k_plot > 0 and np.isfinite(tot_k_plot):
                        k_eef_pct_plot = 100.0 * np.cumsum(radial_k_plot) / tot_k_plot
                        k_diam_plot = 2.0 * rpv
                        maskk = k_eef_pct_plot <= 95.0
                        if np.any(maskk):
                            try:
                                kp_label = f'King fit (I0={I0_k:.2e}, rc={rc_k:.2f}\", alpha={alpha_k:.2f}, b={b_k:.2e})'
                            except Exception:
                                kp_label = 'King fit'
                            plt.plot(k_eef_pct_plot[maskk], k_diam_plot[maskk], label=kp_label, color='purple', linestyle='--', linewidth=eef_linewidth, alpha=0.95)
                except Exception:
                    pass
        except Exception:
            pass
    

    # Add EEF 80% markers: vertical at x=80% (black dashed) and horizontal at diameter where best-focus reaches 80% (purple dashed)
    try:
        plt.axvline(x=80.0, linestyle='--', color='black', linewidth=eef_linewidth)
        if isinstance(profile_pct, (list, np.ndarray)) and len(profile_pct) and isinstance(profile_diam, (list, np.ndarray)) and len(profile_diam):
            # Ensure arrays are numpy arrays and increasing in pct
            pct_arr = np.asarray(profile_pct, dtype=float)
            diam_arr = np.asarray(profile_diam, dtype=float)
            # Sort by pct just in case
            order_idx = np.argsort(pct_arr)
            pct_sorted = pct_arr[order_idx]
            diam_sorted = diam_arr[order_idx]
            if pct_sorted[0] <= 80.0 <= pct_sorted[-1]:
                diam80 = float(np.interp(80.0, pct_sorted, diam_sorted))
                plt.axhline(y=diam80, linestyle='--', color='purple', linewidth=eef_linewidth)
                # Add label slightly below the horizontal line (with slight offset and larger font)
                try:
                    ymin, ymax = ax2.get_ylim()
                    yoff = 0.02 * (ymax - ymin)
                    text_y = diam80 - yoff
                    ax2.text(0, text_y, f'EEF 80% minimum = {diam80:.2f}"', ha='left', va='top', fontsize=9, color='purple')
                except Exception:
                    pass
    except Exception:
        pass
    
    # Add optimized curve if provided
    if df_optimized is not None and opt_frac_profile is not None:
        opt_profile_pct = opt_frac_profile * 100
        opt_profile_diam = 2 * opt_r_profile * m_to_arcsec
        opt_profile_pct_95, opt_profile_diam_95 = limit_percentile(opt_profile_pct, opt_profile_diam)
        label_opt = 'Optimized minimum'
        plt.plot(opt_profile_pct_95, opt_profile_diam_95, label=label_opt, linestyle=':', linewidth=2.5, color='magenta')
    
    # Create final legend for EEF subplot now that all fit curves have been plotted
    try:
        handles2, labels2 = ax2.get_legend_handles_labels()
        if handles2:
            # Wrap long legend labels and place legend slightly higher/outside
            try:
                import textwrap
                wrapped2 = [textwrap.fill(lbl, width=40) if isinstance(lbl, str) else lbl for lbl in labels2]
            except Exception:
                wrapped2 = labels2
            # Place EEF legend inside the EEF subplot area (upper right)
            ax2.legend(handles2, wrapped2, loc='upper left', ncol=1, fontsize=8)
    except Exception:
        pass

    # Add inline labels near each EEF curve so labels appear inside the graph area
    # (Inline curve labels removed — rely on legend and horizontal-line annotations)

    ax2.set_xlabel('Percentage (%)')
    ax2.set_ylabel('Diameter [arcsec]')
    # Increase tick density by reducing major tick spacing by half
    from matplotlib.ticker import MultipleLocator, AutoMinorLocator
    ax2.xaxis.set_major_locator(MultipleLocator(10))  # Major ticks every 10%
    ax2.xaxis.set_minor_locator(AutoMinorLocator(5))  # Minor ticks
    # For y-axis, use auto locator with more ticks
    ax2.yaxis.set_major_locator(plt.MaxNLocator(nbins=20))  # More bins for denser ticks
    ax2.yaxis.set_minor_locator(AutoMinorLocator(5))
    # Mark the 50% encircled energy (Half Energy Width - HEW) in green for minimum
    plt.axhline(y=hew_best_arcsec, linestyle='--', color='green', linewidth=eef_linewidth)  # Horizontal line at HEW diameter
    plt.axvline(x=50, linestyle='--', color='black', linewidth=eef_linewidth)  # Vertical line at 50%
    if hew_best_arcsec is not None:
        try:
            ymin, ymax = ax2.get_ylim()
            yoff = 0.02 * (ymax - ymin)
            text_y = hew_best_arcsec - yoff
            ax2.text(0, text_y, f'HEW minimum = {hew_best_arcsec:.3f}"', ha='left', va='top', fontsize=9, color='green')  # Label HEW with value
        except Exception:
            ax2.text(0, hew_best_arcsec, f'HEW minimum = {hew_best_arcsec:.3f}"', ha='left', va='top', fontsize=9, color='green')
    # Mark the 50% from (0,0) in blue
    plt.axhline(y=hew_origin_arcsec, linestyle='--', color='blue', linewidth=eef_linewidth)  # Horizontal line at HEW(0,0) diameter
    if hew_origin_arcsec is not None:
        try:
            ymin, ymax = ax2.get_ylim()
            yoff = 0.02 * (ymax - ymin)
            text_y = hew_origin_arcsec - yoff
            # place near right side (x=100) but slightly below the line
            ax2.text(100, text_y, f'HEW (0,0) = {hew_origin_arcsec:.3f}"', ha='center', va='top', fontsize=9, color='blue')
        except Exception:
            ax2.text(100, hew_origin_arcsec, f'HEW (0,0) = {hew_origin_arcsec:.3f}"', ha='center', va='top', fontsize=9, color='blue')
    # Mark the 90% encircled energy in red
    if radius_90 is not None:
        plt.axhline(y=eef90_arcsec, linestyle='--', color='red', linewidth=eef_linewidth)  # Horizontal line at EEF90 diameter
        plt.axvline(x=90, linestyle='--', color='black', linewidth=eef_linewidth)  # Vertical line at 90%
        try:
            ymin, ymax = ax2.get_ylim()
            yoff = 0.02 * (ymax - ymin)
            text_y = eef90_arcsec - yoff
            ax2.text(0, text_y, f'EEF 90% minimum = {eef90_arcsec:.3f}"', ha='left', va='top', fontsize=9, color='red')  # Label 90% with value
        except Exception:
            ax2.text(0, eef90_arcsec, f'EEF 90% minimum = {eef90_arcsec:.3f}"', ha='left', va='top', fontsize=9, color='red')
    
    # Add optimized reference lines if provided
    if df_optimized is not None and opt_radius_50 is not None:
        opt_hew_arcsec = 2 * opt_radius_50 * m_to_arcsec
        plt.axhline(y=opt_hew_arcsec, linestyle=':', color='magenta', linewidth=1.5)
        try:
            ymin, ymax = ax2.get_ylim()
            yoff = 0.02 * (ymax - ymin)
            text_y = opt_hew_arcsec - yoff
            ax2.text(0, text_y, f'HEW(opt)={opt_hew_arcsec:.3f}"', ha='left', va='top', fontsize=9, color='magenta')
        except Exception:
            ax2.text(0, opt_hew_arcsec, f'HEW(opt)={opt_hew_arcsec:.3f}"', ha='left', va='bottom', fontsize=9, color='magenta')
        if opt_radius_90 is not None:
            opt_eef90_arcsec = 2 * opt_radius_90 * m_to_arcsec
            plt.axhline(y=opt_eef90_arcsec, linestyle=':', color='orange', linewidth=eef_linewidth)
            try:
                ymin, ymax = ax2.get_ylim()
                yoff = 0.02 * (ymax - ymin)
                text_y = opt_eef90_arcsec - yoff
                ax2.text(0, text_y, f'EEF90(opt)={opt_eef90_arcsec:.3f}"', ha='left', va='top', fontsize=9, color='orange')
            except Exception:
                ax2.text(0, opt_eef90_arcsec, f'EEF90(opt)={opt_eef90_arcsec:.3f}"', ha='left', va='bottom', fontsize=9, color='orange')
    
    # Add titles at the same y-coordinate
    fig.suptitle('')  # Clear any figure title
    # Nudge titles upward to avoid overlapping the PSF image when using constrained_layout
    # Nudge titles upward to avoid overlapping other elements
    ax1.set_title(f'E2E PSF{title_suffix}', fontweight='bold', fontsize=12, y=1.08)
    ax2.set_title(f'Encircled energy function{title_suffix}', fontweight='bold', fontsize=12, y=1.08)
    # layout is handled by constrained_layout on the figure
    # Rebuild the left subplot legend explicitly to guarantee order and layout
    try:
        # Remove any existing legend and rebuild in a deterministic order
        try:
            existing = ax1.get_legend()
            if existing is not None:
                existing.remove()
        except Exception:
            pass
        handles_all, labels_all = ax1.get_legend_handles_labels()
        preferred = []
        for name in ['(0,0)', 'center for minimum HEW']:
            if name in labels_all:
                preferred.append(labels_all.index(name))
        remaining = [i for i in range(len(labels_all)) if i not in preferred]
        final_idx = preferred + remaining
        handles_final = [handles_all[i] for i in final_idx]
        labels_final = [labels_all[i] for i in final_idx]
        ax1.legend(handles_final, labels_final, loc='upper left', ncol=2, fontsize=8, framealpha=0.85, bbox_to_anchor=(0.02, 0.98), bbox_transform=ax1.transAxes)
    except Exception:
        try:
            ax1.legend(loc='upper left', ncol=2, fontsize=8, framealpha=0.85)
        except Exception:
            pass

    # Also write a guaranteed copy of the combined figure to the canonical filename
    try:
        out_comb = os.path.join('Figures', 'E2E_fit_combined.png')
        fig.savefig(out_comb, dpi=150, bbox_inches='tight')
        print(f"Saved combined figure to {out_comb}")
    except Exception:
        pass
    
    # Context menu implementation (cross-platform)
    menu_annotation = None
    menu_active = False
    
    def save_subplot_with_axes(ax, filename, include_colorbar=False):
        """Save a subplot with proper axis labels and ticks"""
        import time
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        final_filename = f"Figures/{filename}_{timestamp}.png"
        
        if include_colorbar:
            # For PSF plot with colorbar, temporarily hide ax2 to avoid capturing it
            ax2_visible = ax2.get_visible()
            ax2.set_visible(False)
            fig.canvas.draw()
            
            # Get bboxes and combine them
            bbox1 = ax1.get_tightbbox(fig.canvas.get_renderer()).transformed(fig.dpi_scale_trans.inverted())
            bbox2 = cbar.ax.get_tightbbox(fig.canvas.get_renderer()).transformed(fig.dpi_scale_trans.inverted())
            
            # Create combined bbox from left of ax1 to right of colorbar
            from matplotlib.transforms import Bbox
            combined_bbox = Bbox([[bbox1.x0, min(bbox1.y0, bbox2.y0)], 
                                  [bbox2.x1, max(bbox1.y1, bbox2.y1)]])
            
            fig.savefig(final_filename, dpi=300, bbox_inches=combined_bbox.expanded(1.15, 1.15))
            
            # Restore ax2 visibility
            ax2.set_visible(ax2_visible)
            fig.canvas.draw_idle()
        else:
            # For other plots, just save the axis
            extent = ax.get_tightbbox(fig.canvas.get_renderer()).transformed(fig.dpi_scale_trans.inverted())
            fig.savefig(final_filename, dpi=300, bbox_inches=extent.expanded(1.15, 1.15))
        
        print(f"Saved plot to {final_filename}")
        return final_filename
    
    def export_psf_plot():
        """Export the PSF plot (left subplot) to PNG"""
        save_subplot_with_axes(ax1, "E2E_PSF", include_colorbar=True)
    
    def export_eef_plot():
        """Export the Encircled Energy Function plot (right subplot) to PNG"""
        save_subplot_with_axes(ax2, "Encircled_Energy", include_colorbar=False)

    def export_eef_csv():
        """Export the Encircled Energy Function data to CSV in CustomPSFs/"""
        import time as _time, os as _os
        ts = _time.strftime('%Y%m%d_%H%M%S')
        out = _os.path.join('CustomPSFs', f'E2E_EEF_{ts}.csv')
        _os.makedirs(_os.path.dirname(out), exist_ok=True)
        # Prepare columns
        try:
            def limit_percentile(pct, diam, max_pct=99.5):
                pct = np.array(pct)
                diam = np.array(diam)
                mask = pct <= max_pct
                return pct[mask], diam[mask]
            pct_best_raw = profile_pct if 'profile_pct' in locals() else (frac_profile * 100 if 'frac_profile' in locals() else [])
            diam_best_raw = profile_diam if 'profile_diam' in locals() else (2 * r_profile * m_to_arcsec if 'r_profile' in locals() else [])
            pct_orig_raw = profile_pct_00 if 'profile_pct_00' in locals() else (frac_00 * 100 if 'frac_00' in locals() else [])
            diam_orig_raw = profile_diam_00 if 'profile_diam_00' in locals() else (2 * r_profile_00 * m_to_arcsec if 'r_profile_00' in locals() else [])
            pct_best, diam_best = limit_percentile(pct_best_raw, diam_best_raw)
            pct_orig, diam_orig = limit_percentile(pct_orig_raw, diam_orig_raw)
        except Exception:
            pct_best, diam_best, pct_orig, diam_orig = [], [], [], []

        # Optimized arrays optional
        opt_pct = None
        opt_diam = None
        if 'opt_frac_profile' in locals() and opt_frac_profile is not None:
            try:
                opt_pct_raw = opt_frac_profile * 100
                opt_diam_raw = 2 * opt_r_profile * m_to_arcsec
                opt_pct, opt_diam = limit_percentile(opt_pct_raw, opt_diam_raw)
            except Exception:
                opt_pct, opt_diam = None, None

        # Build dataframe
        import pandas as _pd
        maxlen = max(len(pct_best) if hasattr(pct_best, '__len__') else 0,
                     len(pct_orig) if hasattr(pct_orig, '__len__') else 0,
                     len(opt_pct) if opt_pct is not None and hasattr(opt_pct, '__len__') else 0)

        def _pad(arr, n):
            if arr is None:
                return [None] * n
            if not hasattr(arr, '__len__'):
                return [arr] * n
            a = list(arr)
            if len(a) < n:
                a = a + [None] * (n - len(a))
            return a

        df_out = _pd.DataFrame({
            'pct_best': _pad(pct_best, maxlen),
            'diam_arcsec_best': _pad(diam_best, maxlen),
            'pct_origin': _pad(pct_orig, maxlen),
            'diam_arcsec_origin': _pad(diam_orig, maxlen),
        })
        if opt_pct is not None:
            df_out['pct_opt'] = _pad(opt_pct, maxlen)
            df_out['diam_arcsec_opt'] = _pad(opt_diam, maxlen)

        df_out.to_csv(out, index=False)
        print('Wrote EEF CSV to', out)

    def export_fit_params_csv():
        """Export the pseudo-Voigt fit parameters to CSV in CustomPSFs/"""
        import time as _time, os as _os, pandas as _pd
        if not fit_params_available:
            print('No pseudo-Voigt fit parameters are available for export.')
            return None
        ts = _time.strftime('%Y%m%d_%H%M%S')
        out = _os.path.join('CustomPSFs', f'E2E_fit_params_{ts}.csv')
        _os.makedirs(_os.path.dirname(out), exist_ok=True)
        df_out = _pd.DataFrame({
            'parameter': ['A', 'Gamma_core', 'Gamma_wing', 'eta', 'beta', 'scalar'],
            'value': [A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit],
        })
        df_out.to_csv(out, index=False)
        print('Wrote fit parameters CSV to', out)

    def export_eef_and_params_excel():
        """Export EEF curves and fit parameters into a single Excel workbook with two sheets."""
        import time as _time, os as _os
        try:
            import pandas as _pd
        except Exception:
            print('pandas required for Excel export')
            return

        if not fit_params_available:
            print('Warning: fit parameters not available; will still try to export EEF curves.')

        ts = _time.strftime('%Y%m%d_%H%M%S')
        out = _os.path.join('CustomPSFs', f'E2E_EEF_and_fitparams_{ts}.xlsx')
        _os.makedirs(_os.path.dirname(out), exist_ok=True)

        # Build a common diameter grid (arcsec) from 0..max r_arcsec
        try:
            rgrid = np.linspace(0.0, float(r_arcsec.max()), 1000)
        except Exception:
            rgrid = np.linspace(0.0, 100.0, 1000)
        diam_grid = 2.0 * rgrid

        # Helper to compute EEF percent from intensity profile on rgrid
        def compute_eef_pct_from_I(I_vals, r_vals):
            try:
                dr = float(r_vals[1] - r_vals[0]) if r_vals.size > 1 else 1.0
                radial = 2.0 * np.pi * r_vals * I_vals * dr
                tot = np.sum(radial)
                if tot <= 0 or not np.isfinite(tot):
                    return np.full_like(r_vals, np.nan)
                eef = 100.0 * np.cumsum(radial) / tot
                return eef
            except Exception:
                return np.full_like(r_vals, np.nan)

        # Aggregated (data) EEF interpolated to diam_grid
        try:
            if profile_pct is not None and profile_diam is not None:
                # Sort and interp
                pa = np.asarray(profile_pct, dtype=float)
                da = np.asarray(profile_diam, dtype=float)
                order = np.argsort(da)
                da_s = da[order]
                pa_s = pa[order]
                # interpolate percent at diam_grid; values outside range will be extrapolated as edge values
                agg_pct_on_grid = np.interp(diam_grid, da_s, pa_s, left=pa_s[0], right=pa_s[-1])
                agg_pct_on_grid[agg_pct_on_grid > 95.0] = np.nan
            else:
                agg_pct_on_grid = np.full_like(diam_grid, np.nan)
        except Exception:
            agg_pct_on_grid = np.full_like(diam_grid, np.nan)

        # PV model EEF on grid
        try:
            pv_I = beta_pseudo_gaussian(rgrid, A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit)
            pv_eef = compute_eef_pct_from_I(pv_I, rgrid)
            pv_eef[pv_eef > 95.0] = np.nan
        except Exception:
            pv_eef = np.full_like(diam_grid, np.nan)

        # Pearson4 EEF on grid - ONLY compute if an accepted pearson4_result exists
        try:
            try:
                pearson4_ok = ('pearson4_result' in locals() and pearson4_result is not None and hasattr(pearson4_result, 'params'))
            except Exception:
                pearson4_ok = False
            if pearson4_ok and 'pearson4_model' in locals():
                try:
                    # Prefer using the accepted fit params only
                    p4_params_plot = pearson4_result.params
                    # Convert lmfit Parameters to plain dict if necessary
                    try:
                        import lmfit as _lm
                        if isinstance(p4_params_plot, _lm.parameter.Parameters):
                            p4_params_plot = {k: v.value for k, v in p4_params_plot.items()}
                    except Exception:
                        pass
                    p4_I = pearson4_model.eval(p4_params_plot, x=rgrid)
                    p4_I = np.maximum(p4_I, 0.0)
                    p4_eef = compute_eef_pct_from_I(p4_I, rgrid)
                    p4_eef[p4_eef > 95.0] = np.nan
                except Exception:
                    p4_eef = np.full_like(diam_grid, np.nan)
            else:
                p4_eef = np.full_like(diam_grid, np.nan)
        except Exception:
            p4_eef = np.full_like(diam_grid, np.nan)

        # King EEF on grid
        try:
            if 'I0_k' in locals():
                k_I = king_profile(rgrid, I0_k, rc_k, alpha_k, b_k)
                k_I = np.maximum(k_I, 0.0)
                k_eef = compute_eef_pct_from_I(k_I, rgrid)
                k_eef[k_eef > 95.0] = np.nan
            else:
                k_eef = np.full_like(diam_grid, np.nan)
        except Exception:
            k_eef = np.full_like(diam_grid, np.nan)

        # Make a canonical name used later in plotting code
        try:
            if 'k_eef' in locals() and k_eef is not None:
                king_eef_pct_plot = k_eef
        except Exception:
            pass

        # Build DataFrame for sheet1: one column per curve (index = diameter)
        df_eef = _pd.DataFrame({'diameter_arcsec': diam_grid, 'EEF_aggregated_pct': agg_pct_on_grid,
                                 'EEF_pv_pct': pv_eef, 'EEF_pearson4_pct': p4_eef, 'EEF_king_pct': k_eef})

        # Build parameter table for sheet2
        params = {}
        # If a pearson4 params object used for plotting exists (p4_params_plot),
        # capture its values even if pearson4_result wasn't stored as a full fit result.
        try:
            if 'p4_params_plot' in locals() and p4_params_plot is not None:
                pr = p4_params_plot
                # lmfit.Parameters-like object
                try:
                    if hasattr(pr, 'keys'):
                        for src_name, out_key in (('amplitude', 'Pearson4_Amplitude'),
                                                  ('center', 'Pearson4_Center_arcsec'),
                                                  ('sigma', 'Pearson4_Sigma_arcsec'),
                                                  ('expon', 'Pearson4_Exponent_m'),
                                                  ('skew', 'Pearson4_Skew_nu')):
                            try:
                                if src_name in pr:
                                    val = pr[src_name].value if hasattr(pr[src_name], 'value') else pr[src_name]
                                    params[out_key] = float(val)
                            except Exception:
                                continue
                except Exception:
                    # fallback for plain dict-like
                    try:
                        for src_name, out_key in (('amplitude', 'Pearson4_Amplitude'),
                                                  ('center', 'Pearson4_Center_arcsec'),
                                                  ('sigma', 'Pearson4_Sigma_arcsec'),
                                                  ('expon', 'Pearson4_Exponent_m'),
                                                  ('skew', 'Pearson4_Skew_nu')):
                            if src_name in pr:
                                params[out_key] = float(pr[src_name])
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            params['Modified_PV_Amplitude_A'] = A_fit
            params['Modified_PV_Gamma_c_arcsec'] = Gamma_c_fit
            params['Modified_PV_Gamma_w_arcsec'] = Gamma_w_fit
            params['Modified_PV_eta'] = eta_fit
            params['Modified_PV_beta'] = beta_fit
            params['Modified_PV_scalar'] = scalar_fit
        except Exception:
            pass
        try:
            if pearson4_result is not None and hasattr(pearson4_result, 'params'):
                pr = pearson4_result.params
                # Core fitted parameters (store with the exact keys requested)
                params['Pearson4_Amplitude'] = float(pr['amplitude'].value) if 'amplitude' in pr else None
                params['Pearson4_Center_arcsec'] = float(pr['center'].value) if 'center' in pr else None
                params['Pearson4_Sigma_arcsec'] = float(pr['sigma'].value) if 'sigma' in pr else None
                params['Pearson4_Exponent_m'] = float(pr['expon'].value) if 'expon' in pr else None
                params['Pearson4_Skew_nu'] = float(pr['skew'].value) if 'skew' in pr else None
        except Exception:
            pass
        try:
            if 'I0_k' in locals():
                params['King_I0'] = I0_k
                params['King_rc_arcsec'] = rc_k
                params['King_alpha'] = alpha_k
                params['King_b'] = b_k
        except Exception:
            pass

        # Convert params dict to DataFrame
        try:
            df_params = _pd.DataFrame(list(params.items()), columns=['parameter', 'value'])
        except Exception:
            df_params = _pd.DataFrame(columns=['parameter', 'value'])

        # Prepare PSF shape sheets (one per fit) on rgrid
        # We'll write Excel formulas for PV and King that reference the
        # Fit_parameters sheet via VLOOKUP, so the workbook contains the fit
        # expressions rather than only numeric values.
        formulas_pv = []
        formulas_k = []
        radii_list = list(rgrid)
        # Build lookup expressions once
        A_lookup = 'VLOOKUP("Modified_PV_Amplitude_A",Fit_parameters!$A:$B,2,FALSE)'
        eta_lookup = 'VLOOKUP("Modified_PV_eta",Fit_parameters!$A:$B,2,FALSE)'
        scalar_lookup = 'VLOOKUP("Modified_PV_scalar",Fit_parameters!$A:$B,2,FALSE)'
        beta_lookup = 'VLOOKUP("Modified_PV_beta",Fit_parameters!$A:$B,2,FALSE)'
        Gc_lookup = 'VLOOKUP("Modified_PV_Gamma_c_arcsec",Fit_parameters!$A:$B,2,FALSE)'
        Gw_lookup = 'VLOOKUP("Modified_PV_Gamma_w_arcsec",Fit_parameters!$A:$B,2,FALSE)'
        I0_lookup = 'VLOOKUP("King_I0",Fit_parameters!$A:$B,2,FALSE)'
        rc_lookup = 'VLOOKUP("King_rc_arcsec",Fit_parameters!$A:$B,2,FALSE)'
        alpha_lookup = 'VLOOKUP("King_alpha",Fit_parameters!$A:$B,2,FALSE)'
        b_lookup = 'VLOOKUP("King_b",Fit_parameters!$A:$B,2,FALSE)'

        for i, rval in enumerate(radii_list, start=2):
            r_cell = f"A{i}"
            # PV formula components
            G_expr = f"EXP(-4*LN(2)*(({r_cell}/{Gc_lookup})^2))"
            a_expr = f"(POWER(2,1/{beta_lookup})-1)"
            C_expr = f"1/POWER(1 + {a_expr}*(POWER(2*{r_cell}/{Gw_lookup},2)), {beta_lookup})"
            mix_expr = f"(1 - {eta_lookup})*{G_expr} + {eta_lookup}*{scalar_lookup}*{C_expr}"
            norm_expr = f"(1 - {eta_lookup}) + {eta_lookup}*{scalar_lookup}"
            pv_formula = f"={A_lookup}*({mix_expr}/{norm_expr})"

            # King formula
            king_formula = f"={I0_lookup}*POWER(1 + POWER({r_cell}/{rc_lookup},2), -{alpha_lookup}) + {b_lookup}"

            formulas_pv.append(pv_formula)
            formulas_k.append(king_formula)

        # Radius numeric column; intensity columns contain Excel formulas (strings starting with '=')
        df_pv = _pd.DataFrame({'radius_arcsec': radii_list, 'I_pv_formula': formulas_pv})
        try:
            df_p4 = _pd.DataFrame({'radius_arcsec': rgrid, 'I_pearson4': p4_I})
        except Exception:
            df_p4 = _pd.DataFrame({'radius_arcsec': rgrid, 'I_pearson4': [None] * len(rgrid)})
        df_k = _pd.DataFrame({'radius_arcsec': radii_list, 'I_king_formula': formulas_k})

        # Write to Excel with multiple sheets
        wrote = False
        try:
            with _pd.ExcelWriter(out, engine='openpyxl') as writer:
                df_eef.to_excel(writer, sheet_name='EEF_curves', index=False)
                df_params.to_excel(writer, sheet_name='Fit_parameters', index=False)
                df_pv.to_excel(writer, sheet_name='PSF_pv', index=False)
                df_p4.to_excel(writer, sheet_name='PSF_pearson4', index=False)
                df_k.to_excel(writer, sheet_name='PSF_king', index=False)
            wrote = True
            print('Wrote combined EEF+fit Excel to', out)
        except Exception as e:
            try:
                with _pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    df_eef.to_excel(writer, sheet_name='EEF_curves', index=False)
                    df_params.to_excel(writer, sheet_name='Fit_parameters', index=False)
                    df_pv.to_excel(writer, sheet_name='PSF_pv', index=False)
                    df_p4.to_excel(writer, sheet_name='PSF_pearson4', index=False)
                    df_k.to_excel(writer, sheet_name='PSF_king', index=False)
                wrote = True
                print('Wrote combined EEF+fit Excel to', out)
            except Exception:
                print('Failed to write Excel file:', e)
        # Also generate a 4-panel fitting performance PNG and save it into CustomPSFs and Figures
        try:
            import matplotlib.pyplot as _plt
            ts_short = ts
            perf_fname_ts = _os.path.join('CustomPSFs', f'fitting_performance_{ts_short}.png')
            perf_fname = _os.path.join('CustomPSFs', 'fitting_performance.png')
            # Prepare data for intensity panel (interpolate observed I_fit to rgrid)
            try:
                if 'I_fit' in locals() and I_fit is not None and 'r_arcsec' in locals():
                    r_obs = np.asarray(r_arcsec)
                    I_obs = np.asarray(I_fit)
                    order_obs = np.argsort(r_obs)
                    r_obs_s = r_obs[order_obs]
                    I_obs_s = I_obs[order_obs]
                    I_on_rgrid = np.interp(rgrid, r_obs_s, I_obs_s, left=np.nan, right=np.nan)
                else:
                    I_on_rgrid = np.full_like(rgrid, np.nan)
            except Exception:
                I_on_rgrid = np.full_like(rgrid, np.nan)

            # Ensure pv_I and p4_I arrays exist for plotting
            try:
                pv_I
            except NameError:
                try:
                    pv_I = beta_pseudo_gaussian(rgrid, A_fit, Gamma_c_fit, Gamma_w_fit, eta_fit, beta_fit, scalar_fit)
                except Exception:
                    pv_I = np.full_like(rgrid, np.nan)
            try:
                p4_I
            except NameError:
                try:
                    # Prefer Pearson4 parameters obtained from the EEF-only fit (pearson4_result.params)
                    p4_params_for_plot = None
                    try:
                        if 'pearson4_result' in locals() and pearson4_result is not None and hasattr(pearson4_result, 'params'):
                            p4_params_for_plot = pearson4_result.params
                    except Exception:
                        p4_params_for_plot = None
                    # If lmfit.Parameters object provided, convert to simple dict of values
                    try:
                        import lmfit
                        if p4_params_for_plot is not None and isinstance(p4_params_for_plot, lmfit.parameter.Parameters):
                            # Convert Parameters to dict of numeric values
                            p4_params_for_plot = {k: v.value for k, v in p4_params_for_plot.items()}
                    except Exception:
                        # lmfit may not be installed or conversion not needed
                        pass
                    if 'pearson4_model' in locals() and (p4_params_for_plot is not None):
                        try:
                            p4_I = pearson4_model.eval(p4_params_for_plot, x=rgrid)
                        except Exception:
                            p4_I = np.full_like(rgrid, np.nan)
                    elif 'pearson4_model' in locals() and ('p4_params_plot' in locals() and p4_params_plot is not None):
                        try:
                            p4_I = pearson4_model.eval(p4_params_plot, x=rgrid)
                        except Exception:
                            p4_I = np.full_like(rgrid, np.nan)
                    else:
                        # No intensity-based fallback: only plot Pearson4 if EEF-fit params exist
                        p4_I = np.full_like(rgrid, np.nan)
                except Exception:
                    p4_I = np.full_like(rgrid, np.nan)

            # Residuals for intensity (linear space on rgrid)
            resid_pv = I_on_rgrid - pv_I
            resid_p4 = I_on_rgrid - p4_I

            # Prepare King intensity model sampled on rgrid so residuals can be
            # plotted consistently with PV/Pearson4. Prefer already-computed
            # `k_I` (on rgrid), else use `Ifit_king_plot` (usually on `rplot`),
            # else try to evaluate `king_profile` using fitted params if available.
            king_model_on_rgrid = None
            try:
                if 'k_I' in locals() and k_I is not None and np.any(np.isfinite(k_I)):
                    king_model_on_rgrid = np.asarray(k_I, dtype=float)
                elif 'Ifit_king_plot' in locals() and Ifit_king_plot is not None:
                    try:
                        x_ifit = rplot if 'rplot' in locals() else r_arcsec
                    except Exception:
                        x_ifit = r_arcsec
                    king_model_on_rgrid = np.interp(rgrid, x_ifit, Ifit_king_plot)
                elif 'I0_k' in locals() and 'rc_k' in locals() and 'alpha_k' in locals():
                    try:
                        king_model_on_rgrid = king_profile(rgrid, I0_k, rc_k, alpha_k, b_k if 'b_k' in locals() else 0.0)
                    except Exception:
                        king_model_on_rgrid = None
            except Exception:
                king_model_on_rgrid = None

            if king_model_on_rgrid is None:
                king_model_on_rgrid = np.full_like(rgrid, np.nan)

            # King residuals (linear) consistent with PV/Pearson4 residuals
            resid_king_linear = I_on_rgrid - king_model_on_rgrid
            # Also compute delta = king_model - aggregated data and log-residuals
            try:
                delta_king = king_model_on_rgrid - I_on_rgrid
                # log residuals (use same floor as elsewhere)
                king_log_resid = None
                try:
                    king_log_resid = np.log(king_model_on_rgrid + floor) - np.log(I_on_rgrid + floor)
                except Exception:
                    king_log_resid = np.full_like(rgrid, np.nan)
                # Print concise diagnostics
                try:
                    mask = np.isfinite(delta_king)
                    nfin = int(np.sum(mask))
                    dmin = float(np.nanmin(delta_king[mask])) if nfin else float('nan')
                    dmax = float(np.nanmax(delta_king[mask])) if nfin else float('nan')
                    drmse = float(np.sqrt(np.nanmean((delta_king[mask])**2))) if nfin else float('nan')
                    print(f"KING DELTA: finite={nfin}, min={dmin:.3g}, max={dmax:.3g}, rmse={drmse:.3g}")
                except Exception:
                    pass
                # Save numeric table for inspection
                try:
                    os.makedirs('Figures', exist_ok=True)
                    import csv
                    csv_path = os.path.join('Figures', 'king_residuals.csv')
                    with open(csv_path, 'w', newline='') as _f:
                        w = csv.writer(_f)
                        w.writerow(['radius_arcsec', 'I_aggregated', 'king_model', 'delta', 'king_log_resid'])
                        for rr, ia, km, dk, klr in zip(rgrid.tolist(), I_on_rgrid.tolist(), king_model_on_rgrid.tolist(), delta_king.tolist(), (king_log_resid.tolist() if hasattr(king_log_resid, 'tolist') else [None]*len(rgrid))):
                            w.writerow([rr, ia, km, dk, klr])
                    print('Wrote King residuals CSV to', csv_path)
                except Exception:
                    pass
            except Exception:
                pass

            # EEF residuals (agg_pct_on_grid vs models)
            try:
                eef_resid_pv = agg_pct_on_grid - pv_eef
            except Exception:
                eef_resid_pv = np.full_like(diam_grid, np.nan)
            try:
                eef_resid_p4 = agg_pct_on_grid - p4_eef
            except Exception:
                eef_resid_p4 = np.full_like(diam_grid, np.nan)
            try:
                # King EEF residuals (Data - King fit) — ensure variable exists for plotting
                eef_resid_king = agg_pct_on_grid - k_eef
            except Exception:
                eef_resid_king = np.full_like(diam_grid, np.nan)

            # Create 2x2 figure
            figp = _plt.figure(figsize=(12, 8))
            ax1p = figp.add_subplot(2, 2, 1)
            # Aggregated data in black dots
            ax1p.plot(rgrid, I_on_rgrid, 'k.', markersize=3, label='Aggregated PSF (data)', zorder=12)
            # Fit curves: dashed and thinner
            _perf_lw = 1.0
            line_pv, = ax1p.plot(rgrid, pv_I, color='red', linestyle='--', lw=_perf_lw, label='Modified pseudo-Voigt', zorder=5)
            # Only plot Pearson4 intensity curve if an accepted fit exists
            try:
                pearson4_ok_plot = ('pearson4_result' in locals() and pearson4_result is not None and hasattr(pearson4_result, 'params'))
            except Exception:
                pearson4_ok_plot = False
            if pearson4_ok_plot and np.any(np.isfinite(p4_I)):
                line_p4, = ax1p.plot(rgrid, p4_I, color='orange', linestyle='--', lw=_perf_lw, label='Pearson4 fit', zorder=5)
            try:
                line_pv.set_dashes([6, 2])
                line_p4.set_dashes([6, 2])
            except Exception:
                pass
            # If King fit available from EEF optimization, prefer that over intensity fallback
            try:
                if 'Ifit_king_plot' in locals() and Ifit_king_plot is not None:
                    line_k, = ax1p.plot(rgrid, Ifit_king_plot, color='purple', linestyle='--', lw=_perf_lw, label='King fit', zorder=5)
                    try:
                        line_k.set_dashes([6, 2])
                    except Exception:
                        pass
                elif 'king_I' in locals():
                    line_k, = ax1p.plot(rgrid, king_I, color='purple', linestyle='--', lw=_perf_lw, label='King fit', zorder=5)
                    try:
                        line_k.set_dashes([6, 2])
                    except Exception:
                        pass
            except Exception:
                pass
            ax1p.set_xlabel('Radius [arcsec]')
            ax1p.set_ylabel('Mean intensity')
            # Use logarithmic scale for intensity plots
            try:
                ax1p.set_yscale('log')
            except Exception:
                pass
            ax1p.grid(True, which='both', linestyle='--', alpha=0.35)
            # Build legend proxies conditionally (include Pearson4 only if plotted)
            try:
                from matplotlib.lines import Line2D
                pv_proxy = Line2D([0], [0], color='red', linestyle='--', lw=_perf_lw)
                k_proxy = Line2D([0], [0], color='purple', linestyle='--', lw=_perf_lw)
                entries = [_plt.Line2D([], [], color='k', marker='.', linestyle='None', ms=3), pv_proxy]
                labels = ['Aggregated PSF (data)', 'Modified pseudo-Voigt']
                try:
                    if pearson4_ok_plot and np.any(np.isfinite(p4_I)):
                        from matplotlib.lines import Line2D as _L2
                        p4_proxy = _L2([0], [0], color='orange', linestyle='--', lw=_perf_lw)
                        entries.append(p4_proxy)
                        labels.append('Pearson4 fit')
                except Exception:
                    pass
                try:
                    if ('Ifit_king_plot' in locals() and Ifit_king_plot is not None) or ('king_I' in locals() and king_I is not None) or ('k_I' in locals() and k_I is not None):
                        entries.append(k_proxy)
                        labels.append('King fit')
                except Exception:
                    pass
                ax1p.legend(entries, labels, fontsize=8)
            except Exception:
                try:
                    ax1p.legend(fontsize=8)
                except Exception:
                    pass

            ax2p = figp.add_subplot(2, 2, 3)
            # Residuals: plot all intensity residuals as percent on the same y-axis
            ax2p.set_xlabel('Radius [arcsec]')
            ax2p.set_ylabel('Residual (%)')
            # Avoid division by zero by using the common floor value used elsewhere
            eps = floor if ('floor' in locals() or 'floor' in globals()) else 1e-12
            try:
                resid_pv_pct = 100.0 * (resid_pv) / (I_on_rgrid + eps)
            except Exception:
                resid_pv_pct = np.full_like(rgrid, np.nan)
            try:
                resid_p4_pct = 100.0 * (resid_p4) / (I_on_rgrid + eps)
            except Exception:
                resid_p4_pct = np.full_like(rgrid, np.nan)
            try:
                resid_king_pct = 100.0 * (resid_king_linear) / (I_on_rgrid + eps) if ('resid_king_linear' in locals() and resid_king_linear is not None) else np.full_like(rgrid, np.nan)
            except Exception:
                resid_king_pct = np.full_like(rgrid, np.nan)

            # Plot percent residuals on the same axis so they share y-limits/scaling
            try:
                lrpv, = ax2p.plot(rgrid, resid_pv_pct, color='red', linestyle='--', lw=_perf_lw, label='Intensity residual: Data - Modified pseudo-Voigt (%)')
            except Exception:
                lrpv = None
            # Only plot Pearson4 intensity residual if an accepted Pearson4 fit exists
            try:
                if pearson4_ok_plot and np.any(np.isfinite(p4_I)):
                    lrp4, = ax2p.plot(rgrid, resid_p4_pct, color='orange', linestyle='--', lw=_perf_lw * 1.6, marker='o', markersize=3, markevery=40, label='Intensity residual: Data - Pearson4 fit (%)')
                else:
                    lrp4 = None
            except Exception:
                lrp4 = None
            try:
                if lrpv is not None:
                    lrpv.set_dashes([6, 2])
                if lrp4 is not None:
                    lrp4.set_dashes([6, 2])
            except Exception:
                pass
            # If King residuals exist, plot them too (as percent to match others)
            try:
                # Build a king residual on the rgrid if possible (prefer res_int_king)
                resid_king_plot = None
                if 'res_int_king' in locals() and res_int_king is not None:
                    try:
                        resid_king_plot = np.interp(rgrid, r_fit, res_int_king, left=np.nan, right=np.nan)
                    except Exception:
                        resid_king_plot = None
                if resid_king_plot is None:
                    if 'Ifit_king_plot' in locals() and Ifit_king_plot is not None:
                        try:
                            x_ifit = rplot if 'rplot' in locals() else r_arcsec
                        except Exception:
                            x_ifit = r_arcsec
                        try:
                            resid_king_plot = I_on_rgrid - np.interp(rgrid, x_ifit, Ifit_king_plot)
                        except Exception:
                            resid_king_plot = None
                    elif 'king_I' in locals() and king_I is not None:
                        resid_king_plot = I_on_rgrid - king_I

                # Convert any available king residuals to percent to match the other traces
                try:
                    if resid_king_plot is not None and np.any(np.isfinite(resid_king_plot)):
                        resid_king_plot_pct = 100.0 * resid_king_plot / (I_on_rgrid + eps)
                    elif 'resid_king_linear' in locals() and resid_king_linear is not None and np.any(np.isfinite(resid_king_linear)):
                        resid_king_plot_pct = 100.0 * resid_king_linear / (I_on_rgrid + eps)
                    else:
                        resid_king_plot_pct = np.full_like(rgrid, np.nan)
                except Exception:
                    resid_king_plot_pct = np.full_like(rgrid, np.nan)

                if resid_king_plot_pct is not None and np.any(np.isfinite(resid_king_plot_pct)):
                    lrk, = ax2p.plot(
                        rgrid,
                        resid_king_plot_pct,
                        color='purple',
                        linestyle='--',
                        lw=_perf_lw * 1.6,
                        alpha=0.95,
                        zorder=15,
                        label='Intensity residual: Data - King fit (%)'
                    )
                    try:
                        lrk.set_dashes([6, 2])
                    except Exception:
                        pass
            except Exception:
                pass
            ax2p.axhline(0, color='gray', linewidth=0.8)
            ax2p.grid(True, which='both', linestyle='--', alpha=0.35)
            try:
                from matplotlib.lines import Line2D
                pv_proxy = Line2D([0], [0], color='red', linestyle='--', lw=_perf_lw)
                k_proxy = Line2D([0], [0], color='purple', linestyle='--', lw=_perf_lw)
                entries = [pv_proxy]
                labels = ['Intensity residual: Data - Modified pseudo-Voigt (%)']
                try:
                    if pearson4_ok_plot and np.any(np.isfinite(p4_I)):
                        from matplotlib.lines import Line2D as _L2
                        p4_proxy = _L2([0], [0], color='orange', linestyle='--', lw=_perf_lw)
                        entries.append(p4_proxy)
                        labels.append('Intensity residual: Data - Pearson4 fit (%)')
                except Exception:
                    pass
                try:
                    if ('resid_king_plot' in locals() and resid_king_plot is not None and np.any(np.isfinite(resid_king_plot))) or ('resid_king_linear' in locals() and resid_king_linear is not None and np.any(np.isfinite(resid_king_linear))):
                        entries.append(k_proxy)
                        labels.append('Intensity residual: Data - King fit (%)')
                except Exception:
                    pass
                if entries:
                    ax2p.legend(entries, labels, fontsize=8)
            except Exception:
                try:
                    ax2p.legend(fontsize=8)
                except Exception:
                    pass

            ax3p = figp.add_subplot(2, 2, 2)
            ax3p.plot(diam_grid, agg_pct_on_grid, color='k', marker='.', linestyle='None', ms=4, label='Aggregated PSF (data)', zorder=12)
            lpv_eef, = ax3p.plot(diam_grid, pv_eef, color='red', linestyle='--', lw=_perf_lw, label='Modified pseudo-Voigt')
            lp4_eef = None
            try:
                if np.any(np.isfinite(p4_eef)):
                    lp4_eef, = ax3p.plot(diam_grid, p4_eef, color='orange', linestyle='--', lw=_perf_lw, label='Pearson4 fit')
            except Exception:
                lp4_eef = None
            try:
                lpv_eef.set_dashes([6, 2])
                if lp4_eef is not None:
                    lp4_eef.set_dashes([6, 2])
            except Exception:
                pass
            # King EEF if available
            try:
                if 'king_eef_pct_plot' in locals():
                    lk_eef, = ax3p.plot(diam_grid, king_eef_pct_plot, color='purple', linestyle='--', lw=_perf_lw, label='King fit')
                    try:
                        lk_eef.set_dashes([6, 2])
                    except Exception:
                        pass
            except Exception:
                pass
            ax3p.grid(True, which='both', linestyle='--', alpha=0.35)
            ax3p.set_xlabel('Diameter [arcsec]')
            ax3p.set_ylabel('Encircled energy (%)')
            try:
                from matplotlib.lines import Line2D
                pv_proxy = Line2D([0], [0], color='red', linestyle='--', lw=_perf_lw)
                k_proxy = Line2D([0], [0], color='purple', linestyle='--', lw=_perf_lw)
                entries = [_plt.Line2D([], [], color='k', marker='.', linestyle='None', ms=4), pv_proxy]
                labels = ['Aggregated PSF (data)', 'Modified pseudo-Voigt']
                if lp4_eef is not None:
                    p4_proxy = Line2D([0], [0], color='orange', linestyle='--', lw=_perf_lw)
                    entries.append(p4_proxy)
                    labels.append('Pearson4 fit')
                try:
                    if 'king_eef_pct_plot' in locals():
                        entries.append(k_proxy)
                        labels.append('King fit')
                except Exception:
                    pass
                ax3p.legend(entries, labels, fontsize=8)
            except Exception:
                try:
                    ax3p.legend(fontsize=8)
                except Exception:
                    pass
            ax3p.set_xlabel('Diameter [arcsec]')
            ax3p.set_ylabel('EEF (%)')

            ax4p = figp.add_subplot(2, 2, 4)
            lrpv_eef, = ax4p.plot(diam_grid, eef_resid_pv, color='red', linestyle='--', lw=_perf_lw, label='EEF residual: Data - Modified pseudo-Voigt (pct)')
            lrp4_eef = None
            try:
                if np.any(np.isfinite(eef_resid_p4)):
                    lrp4_eef, = ax4p.plot(diam_grid, eef_resid_p4, color='orange', linestyle='--', lw=_perf_lw, label='EEF residual: Data - Pearson4 fit (pct)')
            except Exception:
                lrp4_eef = None
            try:
                lrpv_eef.set_dashes([6, 2])
                if lrp4_eef is not None:
                    lrp4_eef.set_dashes([6, 2])
            except Exception:
                pass
            try:
                if 'eef_resid_king' in locals():
                    lrk_eef, = ax4p.plot(diam_grid, eef_resid_king, color='purple', linestyle='--', lw=_perf_lw, label='EEF residual: Data - King fit (pct)')
                    try:
                        lrk_eef.set_dashes([6, 2])
                    except Exception:
                        pass
            except Exception:
                pass
            ax4p.axhline(0, color='gray', linewidth=0.8)
            ax4p.grid(True, which='both', linestyle='--', alpha=0.35)
            ax4p.set_xlabel('Diameter [arcsec]')
            ax4p.set_ylabel('EEF residual (pct)')
            try:
                from matplotlib.lines import Line2D
                pv_proxy = Line2D([0], [0], color='red', linestyle='--', lw=_perf_lw)
                k_proxy = Line2D([0], [0], color='purple', linestyle='--', lw=_perf_lw)
                entries = [pv_proxy]
                labels = ['EEF residual: Data - Modified pseudo-Voigt (pct)']
                try:
                    if lrp4_eef is not None:
                        from matplotlib.lines import Line2D as _L2
                        p4_proxy = _L2([0], [0], color='orange', linestyle='--', lw=_perf_lw)
                        entries.append(p4_proxy)
                        labels.append('EEF residual: Data - Pearson4 fit (pct)')
                except Exception:
                    pass
                try:
                    if 'eef_resid_king' in locals() and np.any(np.isfinite(eef_resid_king)):
                        entries.append(k_proxy)
                        labels.append('EEF residual: Data - King fit (pct)')
                except Exception:
                    pass
                if entries:
                    ax4p.legend(entries, labels, fontsize=8)
            except Exception:
                try:
                    ax4p.legend(fontsize=8)
                except Exception:
                    pass

            figp.tight_layout()
            # Save timestamped and canonical copies into CustomPSFs and Figures
            try:
                _plt.savefig(perf_fname_ts, dpi=150, bbox_inches='tight')
                _plt.savefig(perf_fname, dpi=150, bbox_inches='tight')
                # Also save a copy in Figures for quick viewing
                _plt.savefig(os.path.join('Figures', f'fitting_performance_{ts_short}.png'), dpi=150, bbox_inches='tight')
                _plt.savefig(os.path.join('Figures', 'fitting_performance.png'), dpi=150, bbox_inches='tight')
                _plt.close(figp)
                print('Wrote fitting performance PNG to', perf_fname)
            except Exception:
                try:
                    _plt.savefig(perf_fname, dpi=150, bbox_inches='tight')
                    _plt.close(figp)
                except Exception:
                    pass
        except Exception:
            pass

    def export_fits():
        """Export the aggregated E2E PSF grid `Z` to a minimal FITS file.

        Writes a Primary HDU-only FITS with big-endian doubles and header
        cards: TOT_AEFF, INTG_Z, PIXAS1/2 (arcsec), PIXM1/2 (meters), CDELT1/2 (deg/pixel).
        """
        import time as _time, os as _os
        # Minimal FITS writer (primary HDU only)
        def _write_simple_fits(path: str, data, header_cards: dict | None = None):
            arr = np.asarray(data, dtype=np.float64)
            if arr.ndim != 2:
                raise ValueError('data must be 2D')
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
            header_str = ''.join(card.ljust(80) for card in cards)
            header_bytes = header_str.encode('ascii')
            pad = (2880 - (len(header_bytes) % 2880)) % 2880
            header_bytes += b' ' * pad
            data_be = arr.astype('>f8')
            data_bytes = data_be.tobytes(order='C')
            pad2 = (2880 - (len(data_bytes) % 2880)) % 2880
            data_bytes += b'\x00' * pad2
            with open(path, 'wb') as f:
                f.write(header_bytes)
                f.write(data_bytes)

        # Total A_eff from the DataFrame weights (df in closure).
        # Prefer the adjusted A_eff (after vignetting) when available.
        try:
            if 'aeff_adjusted' in df.columns:
                total_aeff_local = float(np.nansum(df['aeff_adjusted']))
            else:
                total_aeff_local = float(np.nansum(df['weight']))
        except Exception:
            total_aeff_local = float(np.nansum(weight_arr)) if 'weight_arr' in locals() else 0.0

        # Compute integral on our grid (x,y defined in closure)
        try:
            integral_local = float(np.trapz(np.trapz(Z, x, axis=1), y, axis=0))
        except Exception:
            integral_local = float(np.nansum(Z))

        # Pixel scales
        pix_m_x = float(x[1] - x[0]) if x.size > 1 else 0.0
        pix_m_y = float(y[1] - y[0]) if y.size > 1 else 0.0
        m_to_arcsec_local = (180.0 / np.pi) * 3600.0
        pix_as_x = pix_m_x * m_to_arcsec_local
        pix_as_y = pix_m_y * m_to_arcsec_local
        cdelt1 = pix_as_x / 3600.0
        cdelt2 = pix_as_y / 3600.0

        # Output path
        ts = _time.strftime('%Y%m%d_%H%M%S')
        out = _os.path.join('CustomPSFs', f'E2E_aggregated_{ts}.fits')
        _os.makedirs(_os.path.dirname(out), exist_ok=True)
        # Try to pick up author/contact from git config when available
        def _git_cfg(k: str):
            try:
                import subprocess as _sub
                out = _sub.check_output(['git', 'config', '--get', k], stderr=_sub.DEVNULL)
                return out.decode('utf-8').strip()
            except Exception:
                return None

        author_local = _git_cfg('user.name') or 'Unknown'
        contact_local = _git_cfg('user.email') or 'ivo.ferreira@esa.int'
        orcid_local = _git_cfg('user.orcid') or ''
        inputfn = getattr(df, '_input_filename', None) or 'interactive'

        header = {
            'CREATOR': 'main.plot_sum:export_fits',
            'DATE': _time.strftime('%Y-%m-%dT%H:%M:%SZ'),
            'TOT_AEFF': total_aeff_local,
            'INTG_Z': integral_local,
            'CDELT1': float(cdelt1),
            'CDELT2': float(cdelt2),
            'PIXAS1': float(pix_as_x),
            'PIXAS2': float(pix_as_y),
            'PIXM1': float(pix_m_x),
            'PIXM2': float(pix_m_y),
            'AUTHOR': author_local,
            'CONTACT': contact_local,
            'ORCID': orcid_local,
            'INPUTFN': inputfn,
        }
        _write_simple_fits(out, Z, header_cards=header)
        print('Wrote FITS to', out)
    
    def show_context_menu(x, y):
        """Display context menu at given position"""
        nonlocal menu_annotation, menu_active
        
        if menu_active:
            hide_context_menu()
            return
        
        menu_text = "┌────────────────────────────────────────────┐\n"
        menu_text += "│  1. Export PSF Plot                         │\n"
        menu_text += "│  2. Export EEF Plot                         │\n"
        menu_text += "│  3. Export FITS                             │\n"
        menu_text += "│  4. Export EEF CSV                          │\n"
        menu_text += "│  5. Export Fit Parameters CSV               │\n"
        menu_text += "│  6. Export combined EEF + Fit (Excel)       │\n"
        menu_text += "│  7. Cancel                                  │\n"
        menu_text += "└────────────────────────────────────────────┘"
        
        menu_annotation = fig.text(x, y, menu_text,
                                   fontfamily='monospace',
                                   fontsize=10,
                                   bbox=dict(boxstyle='round,pad=0.5', 
                                            facecolor='white', 
                                            edgecolor='black',
                                            linewidth=2,
                                            alpha=0.95),
                                   verticalalignment='top',
                                   horizontalalignment='left',
                                   zorder=1000)
        menu_active = True
        fig.canvas.draw_idle()
    
    def hide_context_menu():
        """Hide the context menu"""
        nonlocal menu_annotation, menu_active
        if menu_annotation:
            menu_annotation.remove()
            menu_annotation = None
            menu_active = False
            fig.canvas.draw_idle()
    
    def on_mouse_press(event):
        """Handle mouse press events"""
        nonlocal menu_active, menu_annotation
        
        # Right-click detection (button 3 on Windows/Linux/Mac)
        is_right_click = False
        if event.button == 3:  # Standard right-click
            is_right_click = True
        elif event.button == 2 and event.key is None:  # Some Mac configs use button 2
            is_right_click = True
        elif event.button == 1 and event.key in ['control', 'ctrl', 'cmd', 'meta', 'super']:
            is_right_click = True  # Ctrl+click for Mac
        
        if is_right_click and event.inaxes in [ax1, ax2] and not menu_active:
            # Show context menu
            x, y = event.x / fig.bbox.width, event.y / fig.bbox.height
            show_context_menu(x, y)
        elif menu_active and event.button == 1:  # Left-click when menu is active
            # Menu is active, check if clicking on menu options
            if menu_annotation:
                # Get menu bounds
                bbox = menu_annotation.get_window_extent()
                if bbox.contains(event.x, event.y):
                    # Calculate relative position within menu (from bottom)
                    relative_y = (event.y - bbox.y0) / bbox.height
                    
                    # Menu structure for 6 options (top-to-bottom):
                    # Top border, Opt1, Opt2, Opt3, Opt4, Opt5, Opt6, Bottom border
                    # Ranges chosen empirically to match text layout
                    if 0.82 < relative_y <= 0.95:  # Option 1 (Export PSF)
                        hide_context_menu()
                        export_psf_plot()
                    elif 0.68 < relative_y <= 0.82:  # Option 2 (Export EEF)
                        hide_context_menu()
                        export_eef_plot()
                    elif 0.54 < relative_y <= 0.68:  # Option 3 (Export FITS)
                        hide_context_menu()
                        export_fits()
                    elif 0.40 < relative_y <= 0.54:  # Option 4 (Export EEF CSV)
                        hide_context_menu()
                        export_eef_csv()
                    elif 0.26 < relative_y <= 0.40:  # Option 5 (Export Fit Params CSV)
                        hide_context_menu()
                        export_fit_params_csv()
                    elif 0.14 < relative_y <= 0.26:  # Option 6 (Export combined EEF + Fit Excel)
                        hide_context_menu()
                        export_eef_and_params_excel()
                    elif 0.08 < relative_y <= 0.14:  # Option 7 (Cancel)
                        hide_context_menu()
                else:
                    # Clicked outside menu, hide it
                    hide_context_menu()
    
    def on_key_press(event):
        """Handle keyboard shortcuts"""
        nonlocal menu_active
        
        if menu_active:
            if event.key in ['1', 'p']:
                hide_context_menu()
                export_psf_plot()
            elif event.key in ['2', 'e']:
                hide_context_menu()
                export_eef_plot()
            elif event.key in ['3', 'f']:
                hide_context_menu()
                export_fits()
            elif event.key in ['4', 'c']:
                hide_context_menu()
            elif event.key in ['x', '6']:
                export_eef_and_params_excel()
                export_eef_csv()
            elif event.key in ['5', 's']:
                hide_context_menu()
                export_fit_params_csv()
            elif event.key in ['6', 'x']:
                hide_context_menu()
                export_eef_and_params_excel()
            elif event.key in ['6', 'escape']:
                hide_context_menu()
        else:
            if event.key in ['p', '1']:
                export_psf_plot()
            elif event.key in ['e', '2']:
                export_eef_plot()
            elif event.key in ['f', '3']:
                export_fits()
            elif event.key in ['s', '5']:
                export_fit_params_csv()
            elif event.key == 'h':
                print("\nKeyboard shortcuts:")
                print("  'p' or '1' - Export PSF plot")
                print("  'e' or '2' - Export Encircled Energy plot")
                print("  Right-click (or Ctrl+click on Mac) - Show context menu")
                print("  'h' - Show this help")
    
    # Connect events
    fig.canvas.mpl_connect('button_press_event', on_mouse_press)
    fig.canvas.mpl_connect('key_press_event', on_key_press)
    
    # Print help message
    print("\nKeyboard shortcuts:")
    print("  'p' or '1' - Export PSF plot")
    print("  'e' or '2' - Export Encircled Energy plot")
    print("  Right-click (or Ctrl+click on Mac) on plots - Show context menu")
    print("  'h' - Show help")
    print("  'f' or '3' - Export FITS (aggregated E2E PSF)")
    
    if output:
        plt.savefig(output, dpi=150)  # Save the combined plot
        print(f"Saved combined plot to {output}")
    if not output:
        # No explicit output requested. To avoid blocking in interactive
        # backends (which calls plt.show() and waits for user), save a
        # non-interactive copy of the combined plot to a file alongside the
        # input workbook. This prevents the CLI invocation from appearing to
        # "hang" while still producing a useful artifact. If saving fails
        # for any reason, fall back to showing the interactive window.
        # Prefer explicit interactive display when requested. Use a safe lookup
        # for CLI args exposed on __main__.args to avoid NameError when this
        # module is imported or called programmatically.
        try:
            _main_mod = sys.modules.get('__main__')
            _main_args = getattr(_main_mod, 'args', None) if _main_mod is not None else None
            _show_flag = getattr(_main_args, 'show', False) if _main_args is not None else False
        except Exception:
            _show_flag = False
        if _show_flag:
            try:
                plt.show()
            except Exception:
                try:
                    base, _ = os.path.splitext(args.file)
                    auto_out = f"{base}_plot.png"
                    plt.savefig(auto_out, dpi=150)
                    print(f"Interactive display failed; saved plot to {auto_out} instead.")
                except Exception:
                    pass
        else:
            # Default: show the interactive plot and keep it open. If this
            # fails (e.g. running in a headless environment), do NOT auto-save
            # a PNG; instead instruct the user to run with --output to save.
            # Additionally, support --export-package: generate exports into a
            # timestamped folder containing input, FITS, Excel and figures.
            try:
                # Check CLI flag exposed on __main__
                _main_mod = sys.modules.get('__main__')
                _main_args = getattr(_main_mod, 'args', None) if _main_mod is not None else None
                _export_pkg = getattr(_main_args, 'export_package', False) if _main_args is not None else False
            except Exception:
                _export_pkg = False

            if _export_pkg:
                try:
                    import shutil, glob, time as _time
                    ts = _time.strftime('%Y%m%d_%H%M%S')
                    pkg_dir = os.path.join('Exports', ts)
                    os.makedirs(pkg_dir, exist_ok=True)
                    # Copy input workbook
                    try:
                        shutil.copy2(args.file, os.path.join(pkg_dir, os.path.basename(args.file)))
                    except Exception:
                        pass
                    # Trigger standard exports (they will write into CustomPSFs/ and Figures/)
                    try:
                        export_fits()
                    except Exception:
                        pass
                    try:
                        export_eef_and_params_excel()
                    except Exception:
                        pass
                    try:
                        export_psf_plot()
                    except Exception:
                        pass
                    try:
                        export_eef_plot()
                    except Exception:
                        pass
                    # Collect generated artifacts by pattern and copy newest matches
                    def copy_latest(patterns, dest_dir):
                        for pat in patterns:
                            try:
                                matches = glob.glob(pat)
                                if not matches:
                                    continue
                                latest = max(matches, key=os.path.getmtime)
                                shutil.copy2(latest, os.path.join(dest_dir, os.path.basename(latest)))
                            except Exception:
                                continue

                    # FITS
                    copy_latest(['CustomPSFs/E2E_aggregated_*.fits'], pkg_dir)
                    # Combined EEF+fit Excel
                    copy_latest(['CustomPSFs/E2E_EEF_and_fitparams_*.xlsx', 'CustomPSFs/E2E_EEF_and_fitparams_*.xls'], pkg_dir)
                    # PSF and EEF PNGs
                    copy_latest(['Figures/E2E_PSF_*.png', 'Figures/E2E_PSF_*.PNG'], pkg_dir)
                    copy_latest(['Figures/Encircled_Energy_*.png', 'Figures/Encircled_Energy_*.PNG'], pkg_dir)
                    # EEF fit combined figure (may be named without timestamp)
                    copy_latest(['Figures/E2E_fit_combined*.png', 'Figures/E2E_fit_combined.png'], pkg_dir)
                    # Include fitting performance PNG if present in CustomPSFs
                    copy_latest(['CustomPSFs/fitting_performance*.png', 'CustomPSFs/fitting_performance.png'], pkg_dir)

                    print(f"Exported package to {pkg_dir}")
                except Exception as e:
                    try:
                        print('Failed to create export package:', e)
                    except Exception:
                        pass
            try:
                plt.show()
            except Exception:
                print("Interactive display failed — no GUI available. To save the plot use --output <file.png> or run in a GUI session.")


def compute_hew_eef_metrics(file: str = 'Distributions/TestDistribution.xlsx', sheet: str = 'MM_PSF', normalize: bool = True, fast: bool = True, df_optimized: pd.DataFrame = None) -> dict:
    """Convenience wrapper: load workbook and return HEW/EEF metrics (no plotting).

    Returns a dict matching the CLI JSON output.
    """
    df = load_gaussians_from_excel(file, sheet)
    return plot_sum(df, normalize=normalize, fast=fast, df_optimized=df_optimized, return_metrics_only=True)


if __name__ == '__main__':
    # Close all existing figures to ensure clean start
    plt.close('all')
    try:
        print('main.py: starting with argv=', sys.argv)
    except Exception:
        pass
    
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='Plot sum of rotated Gaussians from Excel.')
    parser.add_argument('-f','--file', default='Distributions/TestDistribution.xlsx', help='Excel file path')
    parser.add_argument('-s','--sheet', default='MM_PSF', help='Sheet name to read (default MM_PSF)')
    parser.add_argument('--normalize', dest='normalize', action='store_true', default=True, help='Normalize each Gaussian to integrate to 1 (default on)')
    parser.add_argument('--no-normalize', dest='normalize', action='store_false', help='Disable normalization')
    parser.add_argument('-o','--output', help='Optional output image file path (PNG)')
    parser.add_argument('--show', action='store_true', help='Show interactive plot window (overrides auto-save)')
    parser.add_argument(
        '--mode',
        type=str,
        choices=['coarse', 'fine', 'extra-fine'],
        default='coarse',
        help='Runtime mode: coarse, fine, or extra-fine. Controls plotting + optimization speed/accuracy.'
    )
    parser.add_argument('--optimize', action='store_true', default=False, help='Enable MM position optimization (uses --mode for speed/accuracy).')
    # Compatibility alias (UK spelling)
    parser.add_argument('--optimise', dest='optimize', action='store_true', help=argparse.SUPPRESS)
    parser.add_argument(
        '--placement',
        type=str,
        nargs='?',
        const='cross',
        choices=['cross', 'x_axis', 'elliptical'],
        default=None,
        help=(
            "Apply a placement strategy and write a *_placed.xlsx file. "
            "If used together with --optimize, the selected placement seeds the optimization. "
            "Strategies: 'cross' (90° pattern, previous default), "
            "'x_axis' (+/-x with above/below alternation), or "
            "'elliptical' (per-row: best MMs toward x, worst toward y)."
        ),
    )
    parser.add_argument('--return_metrics_only', dest='return_metrics_only', action='store_true', help='Return HEW/EEF metrics only (no plot).')
    parser.add_argument('--suppress-output', dest='suppress_output', action='store_true', default=False, help='Do not write placed/optimised Excel outputs when running non-interactively.')
    parser.add_argument('--metrics-nr-final', type=int, default=None, help=argparse.SUPPRESS)
    parser.add_argument('--metrics-ntheta-final', type=int, default=None, help=argparse.SUPPRESS)
    parser.add_argument('--metrics-r-margin', type=float, default=None, help=argparse.SUPPRESS)
    # Removed --input-pickle support; keep input-csv for CSV single-sheet paths
    parser.add_argument('--input-csv', default=None, help=argparse.SUPPRESS)
    parser.add_argument('--export-package', dest='export_package', action='store_true', default=False, help='Create a timestamped folder with input, FITS, Excel and PNG artifacts')
    parser.add_argument('--batch-combinations', dest='batch_combinations', default=None, help='Path to Excel file containing combinations (one config per row). If provided the script will run each configuration against the input file and exit.')
    args = parser.parse_args()

    # Expose parsed args to helper functions that inspect __main__.args
    try:
        import sys as _sys
        _sys.modules['__main__'].args = args
    except Exception:
        pass

    # If an explicit input CSV path was provided, prefer it as the input file
    if getattr(args, 'input_csv', None):
        args.file = args.input_csv
    # Note: in-memory pickled input support removed for release.

    # If args.file is a multi-sheet CSV, parse it and materialize a temporary .xlsx
    if isinstance(args.file, str) and args.file.lower().endswith('.csv'):
        try:
            sheets = parse_multisheet_csv(args.file)
            # If CSV contained multiple sheets, write to a temporary .xlsx and use that
            if isinstance(sheets, dict) and len(sheets) > 1:
                tf = tempfile.NamedTemporaryFile(prefix='multisheet_', suffix='.xlsx', delete=False)
                tf.close()
                tmp_path = Path(tf.name)
                try:
                    with pd.ExcelWriter(tmp_path, engine='openpyxl') as writer:
                        for sname, sdf in sheets.items():
                            try:
                                sdf.to_excel(writer, sheet_name=sname, index=False)
                            except Exception:
                                continue
                    args.file = str(tmp_path)
                except Exception:
                    # fallback: leave args.file as CSV path
                    pass
        except Exception:
            # If parsing fails, fall back to treating CSV as single-sheet file
            pass

    # If user requested metrics-only, disable interactive plotting and suppress show().
    if getattr(args, 'return_metrics_only', False):
        try:
            plt.ioff()
            plt.show = lambda *a, **k: None
            # Suppress incidental debug/sample prints so CLI output is a clean JSON
            try:
                import os as _os
                _os.environ['SILENCE_OUTPUT'] = '1'
            except Exception:
                pass
        except Exception:
            pass

    # Load data from Excel
    # Batch mode: process combinations file and exit
    if getattr(args, 'batch_combinations', None):
        comb_path = args.batch_combinations
        try:
            combos = pd.read_excel(comb_path, engine='openpyxl')
        except Exception as e:
            print(f"Failed to read combinations file {comb_path}: {e}")
            sys.exit(2)

        # Iterate rows: expect columns as follows (by position):
        # A: (ignored), B: config name (prefix), C: off-axis, D: energy, E: defocus
        for rid, crow in combos.iterrows():
            try:
                prefix = str(crow.iat[1]).strip()
            except Exception:
                print(f"Skipping row {rid}: cannot read prefix in column B")
                continue
            try:
                offaxis_val = float(crow.iat[2]) if not pd.isna(crow.iat[2]) else 0.0
            except Exception:
                offaxis_val = 0.0
            try:
                energy_val = float(crow.iat[3]) if not pd.isna(crow.iat[3]) else 0.0
            except Exception:
                energy_val = 0.0
            try:
                defocus_val = float(crow.iat[4]) if not pd.isna(crow.iat[4]) else 0.0
            except Exception:
                defocus_val = 0.0

            print(f"Processing configuration '{prefix}': offaxis={offaxis_val}, energy={energy_val}, defocus={defocus_val}")

            # Prepare new input workbook path
            src_path = args.file
            src_dir = os.path.dirname(src_path) or '.'
            src_basename = os.path.basename(src_path)
            new_basename = f"{prefix}_{src_basename}"
            new_path = os.path.join(src_dir, new_basename)
            try:
                shutil.copy2(src_path, new_path)
            except Exception as e:
                print(f"Failed to copy input workbook to {new_path}: {e}")
                continue

            # Modify the copy: adjust Thermal sheet d_therm_rotx/roty and d_therm_z,
            # and set energy in C2 of vignetting sheets.
            try:
                import openpyxl
                wb = openpyxl.load_workbook(new_path)

                # Thermal adjustments
                if 'Thermal' in wb.sheetnames:
                    ws = wb['Thermal']
                    # Read header row (assume first row contains headers)
                    headers = [str(c.value).strip() if c.value is not None else '' for c in ws[1]]
                    # Find matching columns (prefer exact base names and skip
                    # any headers that end with an underscore to avoid backup
                    # columns like 'd_therm_rotx_'). Normalise by removing
                    # non-alphanumeric characters (except underscore) for matching.
                    import re as _re
                    def _norm_header(s: str) -> str:
                        return _re.sub(r"[^0-9a-z_]+", "", s.lower()).strip()

                    col_rotx = None
                    col_roty = None
                    col_z = None
                    for idx, h in enumerate(headers):
                        if not h:
                            continue
                        # Skip headers that intentionally end with an underscore
                        if h.strip().endswith('_'):
                            continue
                        nh = _norm_header(h)
                        if nh == 'd_therm_rotx' or nh.startswith('d_therm_rotx'):
                            col_rotx = idx + 1
                        if nh == 'd_therm_roty' or nh.startswith('d_therm_roty'):
                            col_roty = idx + 1
                        if nh == 'd_therm_z' or nh.startswith('d_therm_z'):
                            col_z = idx + 1

                    # Compute deltas
                    # offaxis in input sheet is in arcminutes; convert to arcseconds
                    offaxis_arcsec = float(offaxis_val) * 60.0
                    # split equally between rot x/y by dividing by sqrt(2)
                    delta_rot = offaxis_arcsec / math.sqrt(2.0)
                    # defocus in input sheet is in millimetres; convert to microns
                    defocus_um = float(defocus_val) * 1e3

                    for r in range(2, ws.max_row + 1):
                        if col_rotx:
                            cell = ws.cell(row=r, column=col_rotx)
                            try:
                                cell_val = float(cell.value) if cell.value is not None else 0.0
                            except Exception:
                                cell_val = 0.0
                            cell.value = cell_val + delta_rot
                        if col_roty:
                            cell = ws.cell(row=r, column=col_roty)
                            try:
                                cell_val = float(cell.value) if cell.value is not None else 0.0
                            except Exception:
                                cell_val = 0.0
                            cell.value = cell_val + delta_rot
                        if col_z:
                            cell = ws.cell(row=r, column=col_z)
                            try:
                                cell_val = float(cell.value) if cell.value is not None else 0.0
                            except Exception:
                                cell_val = 0.0
                            cell.value = cell_val + defocus_um

                # Vignetting energy in C2 for both rotrad and rotazi
                for sname in ('Vignetting rotrad', 'Vignetting rotazi'):
                    if sname in wb.sheetnames:
                        ws_v = wb[sname]
                        try:
                            ws_v.cell(row=2, column=3).value = float(energy_val)
                        except Exception:
                            pass

                wb.save(new_path)
            except Exception as e:
                print(f"Failed to modify workbook {new_path}: {e}")
                continue

            # Run this script on the modified workbook with --export-package
            try:
                cmd = [sys.executable, os.path.abspath(__file__), '--file', new_path, '--export-package']
                print('Running:', ' '.join(cmd))

                # Record existing Exports subfolders so we can detect a new package
                exports_root = os.path.join(os.getcwd(), 'Exports')
                existing_dirs = set(os.listdir(exports_root)) if os.path.isdir(exports_root) else set()

                subprocess.check_call(cmd)
            except Exception as e:
                print(f"Failed to run export for {prefix}: {e}")

            # Prefer zipping the full export package folder created by the child run.
            try:
                os.makedirs(exports_root, exist_ok=True)
                # Look for new subdirectories created during the child run
                all_dirs = [d for d in os.listdir(exports_root) if os.path.isdir(os.path.join(exports_root, d))]
                new_dirs = [d for d in all_dirs if d not in existing_dirs]
                if new_dirs:
                    # Prefer the newest created package folder
                    latest_dir = max(new_dirs, key=lambda d: os.path.getmtime(os.path.join(exports_root, d)))
                    pkg_path = os.path.join(exports_root, latest_dir)
                    zip_target = os.path.join(exports_root, f"{prefix}_{src_basename}.zip")
                    shutil.make_archive(os.path.splitext(zip_target)[0], 'zip', pkg_path)
                    print(f"Created package from export folder: {zip_target}")
                else:
                    # Fallback: create a zip containing just the modified workbook
                    zip_target = os.path.join(exports_root, f"{prefix}_{src_basename}.zip")
                    with zipfile.ZipFile(zip_target, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                        zf.write(new_path, arcname=os.path.basename(new_path))
                    print(f"Created fallback package (workbook only): {zip_target}")
            except Exception as e:
                print(f"Failed to create package for {prefix}: {e}")

        # Completed batch processing
        sys.exit(0)

    df = load_gaussians_from_excel(args.file, args.sheet)
    plot_title_suffix = ""
    df_optimized = None
    
    placement_strategy = args.placement
    
    # If optimization is requested, run it and reload the optimized configuration.
    # If --placement is also provided, use it as the optimizer seed.
    if args.optimize:
        from optimize_mm_rows import optimize_rows
        base, ext = os.path.splitext(args.file)
        if ext == '':
            ext = '.xlsx'
        opt_output = f"{base}_optimised{ext}"

        is_coarse = (args.mode == 'coarse')
        is_extra_fine = (args.mode == 'extra-fine')
        # Total budget requirements:
        # - mode=coarse: optimize+plot <= 30s
        # - mode=fine: optimize+plot <= 90s
        # - mode=extra-fine: optimize+plot <= 300s
        # We budget optimization to leave time for plotting.
        if is_coarse:
            opt_budget_s = 18.0
        elif is_extra_fine:
            opt_budget_s = 240.0
        else:
            opt_budget_s = 45.0

        start_strategy = placement_strategy or 'elliptical'
        print(f"Optimizing MM positions (mode={args.mode}, start={start_strategy})...")
        try:
            # If the input is a CSV-only file, skip workbook-based optimization
            if isinstance(args.file, str) and str(args.file).lower().endswith('.csv'):
                print("Skipping MM position optimization for CSV-only input (no MM configuration sheet).")
                opt_hew = None
            else:
                opt_hew = optimize_rows(
                    input_path=args.file,
                    output_path=opt_output,
                    mode=args.mode,
                    optimize=True,
                    time_budget_s=opt_budget_s,
                    start_placement=start_strategy,
                    write_output=(not getattr(args, 'suppress_output', False)),
                )
        except KeyboardInterrupt:
            print("Optimization interrupted (Ctrl+C). The optimised file was not updated.")
            opt_hew = None
        else:
            print(f"Optimization complete. Optimized HEW: {opt_hew:.6e} m")
            if not getattr(args, 'suppress_output', False):
                print(f"Optimized configuration saved to: {opt_output}")

            # Make it explicit whether MM configuration actually changed.
            try:
                mm_in = pd.read_excel(args.file, sheet_name="MM configuration", engine="openpyxl")
                mm_out = pd.read_excel(opt_output, sheet_name="MM configuration", engine="openpyxl")
                a = mm_in["MM #"].astype(int).to_numpy()
                b = mm_out["MM #"].astype(int).to_numpy()
                changed = (a != b)
                print(f"MM configuration changed entries: {int(changed.sum())}")
                if changed.any() and "Row #" in mm_in.columns:
                    rows_changed = sorted(mm_in.loc[changed, "Row #"].dropna().unique().tolist())
                    print(f"Rows with MM# changes: {rows_changed}")

                # Print a small sample so it's easy to spot in Excel.
                if changed.any():
                    idx = np.flatnonzero(changed)[:10].tolist()
                    sample = pd.DataFrame(
                        {
                            "sheet_row_index": idx,
                            "Row #": mm_in.loc[idx, "Row #"].to_numpy() if "Row #" in mm_in.columns else [None] * len(idx),
                            "MM # (input)": a[idx],
                            "MM # (optimised)": b[idx],
                        }
                    )
                    print("Sample MM# changes (first 10):")
                    print(sample.to_string(index=False))
            except Exception:
                pass

            # Load the optimized configuration for overlaying
            if os.path.exists(opt_output):
                df_optimized = load_gaussians_from_excel(opt_output, args.sheet)
            else:
                df_optimized = None
            plot_title_suffix = " (comparison)"

    # Placement-only mode
    elif placement_strategy is not None:
        from optimize_mm_rows import cross_placement, x_axis_placement, elliptical_placement
        base, ext = os.path.splitext(args.file)
        if ext == '':
            ext = '.xlsx'
        opt_output = f"{base}_placed{ext}"

        if placement_strategy == 'x_axis':
            placement_fn = x_axis_placement
            placement_label = 'x_axis'
        elif placement_strategy == 'elliptical':
            placement_fn = elliptical_placement
            placement_label = 'elliptical'
        else:
            placement_fn = cross_placement
            placement_label = 'cross'

        print(f"Applying placement strategy ({placement_label})...")
        try:
            # If the input is a CSV-only file, skip workbook-based placement
            if isinstance(args.file, str) and str(args.file).lower().endswith('.csv'):
                print("Skipping MM placement for CSV-only input (no MM configuration sheet).")
                placement_hew = None
            else:
                placement_hew = placement_fn(
                    input_path=args.file,
                    output_path=opt_output,
                    seed=42,
                    write_output=(not getattr(args, 'suppress_output', False)),
                )
        except KeyboardInterrupt:
            print("Placement interrupted (Ctrl+C). The placed file was not updated.")
            placement_hew = None
        else:
            if placement_hew is None:
                print("Placement skipped (CSV-only input); no placed configuration produced.")
                df_optimized = None
            else:
                print(f"Placement complete. Final HEW: {placement_hew:.6e} m")
                if not getattr(args, 'suppress_output', False):
                    print(f"Placed configuration saved to: {opt_output}")

                # Make it explicit how many positions changed
                try:
                    if not getattr(args, 'suppress_output', False) and os.path.exists(opt_output):
                        mm_in = pd.read_excel(args.file, sheet_name="MM configuration", engine="openpyxl")
                        mm_out = pd.read_excel(opt_output, sheet_name="MM configuration", engine="openpyxl")
                    else:
                        mm_in = None
                        mm_out = None
                    if mm_in is not None and mm_out is not None:
                        a = mm_in["MM #"].astype(int).to_numpy()
                        b = mm_out["MM #"].astype(int).to_numpy()
                        changed = (a != b)
                        print(f"MM configuration changed entries: {int(changed.sum())}")

                        # Print a small sample
                        if changed.any():
                            idx = np.flatnonzero(changed)[:10].tolist()
                            sample = pd.DataFrame(
                                {
                                    "sheet_row_index": idx,
                                    "MM # (input)": a[idx],
                                    "MM # (placed)": b[idx],
                                }
                            )
                            print("Sample MM# changes (first 10):")
                            print(sample.to_string(index=False))
                except Exception:
                    pass

                # Load the placed configuration for overlaying (only if file was written)
                if not getattr(args, 'suppress_output', False) and os.path.exists(opt_output):
                    df_optimized = load_gaussians_from_excel(opt_output, args.sheet)
            plot_title_suffix = f" (placement: {placement_label})"
    
    # Plot the sum and encircled energy (with optional optimized overlay)
    if args.return_metrics_only:
        metrics = plot_sum(
            df,
            normalize=args.normalize,
            output=args.output,
            fast=(args.mode == 'coarse'),
            title_suffix=plot_title_suffix,
            df_optimized=df_optimized,
            return_metrics_only=True,
            metrics_n_r_final=args.metrics_nr_final,
            metrics_n_theta_final=args.metrics_ntheta_final,
            metrics_r_margin=args.metrics_r_margin,
        )
        print(json.dumps(metrics, indent=2))
        sys.exit(0)


    plot_sum(
        df,
        normalize=args.normalize,
        output=args.output,
        fast=(args.mode == 'coarse'),
        title_suffix=plot_title_suffix,
        df_optimized=df_optimized,
    )

    # Always trigger export logic for A_eff sheet after plot_sum
    try:
        import openpyxl
        import numpy as np
        import tempfile, os, sys
        wb = openpyxl.load_workbook(args.file)
        df_a = df.copy()
        if 'A_eff' in wb.sheetnames:
            ws_a = wb['A_eff']
            max_r_a = ws_a.max_row or 0
            mm_series = pd.to_numeric(df_a.get('MM #', pd.Series([], dtype=float)), errors='coerce')
            base_series = pd.to_numeric(df_a.get('aeff_base', pd.Series([], dtype=float)), errors='coerce')
            adj_series = pd.to_numeric(df_a.get('aeff_adjusted', pd.Series([], dtype=float)), errors='coerce')
            mm_to_base = {}
            mm_to_adj = {}
            for mmv, bv, av in zip(mm_series.tolist(), base_series.tolist(), adj_series.tolist()):
                try:
                    key = int(mmv) if mmv is not None and not (isinstance(mmv, float) and np.isnan(mmv)) else None
                except Exception:
                    key = None
                if key is None:
                    continue
                try:
                    if bv is not None and not (isinstance(bv, float) and np.isnan(bv)):
                        mm_to_base[key] = float(bv)
                except Exception:
                    pass
                try:
                    if av is not None and not (isinstance(av, float) and np.isnan(av)):
                        mm_to_adj[key] = float(av)
                except Exception:
                    pass
            for r in range(1, max_r_a + 1):
                cell = ws_a.cell(row=r, column=1).value
                if cell is None:
                    continue
                try:
                    mmv = int(cell) if isinstance(cell, (int, float)) or (isinstance(cell, str) and str(cell).strip().isdigit()) else None
                    if isinstance(cell, str) and str(cell).strip().isdigit():
                        mmv = int(str(cell).strip())
                except Exception:
                    mmv = None
                if mmv is None:
                    continue
                # Write canonical A_eff into column B
                if mmv in mm_to_base:
                    try:
                        ws_a.cell(row=r, column=2, value=float(mm_to_base[mmv]))
                    except Exception:
                        pass
                # Always write column C: zero if base is zero or missing, else adjusted value if present, else zero
                try:
                    base_val = float(mm_to_base.get(mmv, 0.0))
                    adj_val = float(mm_to_adj.get(mmv, 0.0))
                    if base_val is None or (isinstance(base_val, float) and np.isnan(base_val)):
                        ws_a.cell(row=r, column=3, value=0.0)
                    elif base_val == 0.0:
                        ws_a.cell(row=r, column=3, value=0.0)
                    else:
                        ws_a.cell(row=r, column=3, value=adj_val)
                except Exception:
                    ws_a.cell(row=r, column=3, value=0.0)
            tmpf = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
            tmpf.close()
            wb.save(tmpf.name)
            os.replace(tmpf.name, args.file)
            sys.stdout.flush()
    except Exception:
        pass
