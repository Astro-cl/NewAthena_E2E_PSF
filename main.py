import argparse  # For parsing command-line arguments
import numpy as np  # For numerical operations and arrays
import pandas as pd  # For data manipulation and Excel reading
import matplotlib
import matplotlib.pyplot as plt  # For plotting
import matplotlib.gridspec as gridspec
import os
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
def load_aeff_weight_map(path: str) -> dict:
    """Load A_eff mapping (MM # -> base A_eff) from `A_eff` sheet.

    Reads the `A_eff` sheet (headerless) and returns a dict mapping integer
    MM -> float(A_eff_base). If the sheet is missing or no valid rows are
    found an empty dict is returned.
    """
    mapping = {}
    try:
        raw = pd.read_excel(path, sheet_name='A_eff', engine='openpyxl', header=None)
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
            continue
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
                # Create mapping from MM # to theta_degrees
                # theta_cw_from_y: angle clockwise from y-axis (this is what user specifies)
                # Both theta_position and theta_degrees use the same formula: -(theta_cw_from_y - 90)
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

            # Attempt to detect an adjusted-A_eff column (written into column C
            # by the GUI/exporter) and prefer it for initial weights while
            # preserving `aeff_base` as the bare throughput.
            aeff_map_adjusted = {}
            try:
                raw_a = pd.read_excel(path, sheet_name='A_eff', engine='openpyxl', header=None)
                # Expect layout: col0=MM #, col1=A_eff (base), col2=A_eff_adjusted
                if raw_a.shape[1] >= 3:
                    mm_col = pd.to_numeric(raw_a.iloc[:, 0], errors='coerce')
                    adj_col = pd.to_numeric(raw_a.iloc[:, 2], errors='coerce')
                    valid = mm_col.notna() & adj_col.notna()
                    if valid.any():
                        for mmv, adjv in zip(mm_col[valid].astype(int).tolist(), adj_col[valid].astype(float).tolist()):
                            aeff_map_adjusted[int(mmv)] = float(adjv)
            except Exception:
                aeff_map_adjusted = {}

            # Initialize weight: prefer adjusted map when available, else base.
            if aeff_map_adjusted:
                df['weight'] = mm_as_int.map(aeff_map_adjusted).astype(float)
                # Fill any missing weights from base
                df['weight'] = df['weight'].fillna(df['aeff_base'].astype(float))
            else:
                df['weight'] = df['aeff_base'].astype(float)

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

    # --- Apply polar vignetting (rotazi + rotrad) after A_eff/weight initialization ---
    try:
        # compute rotation projections using the populated mm_to_pos and *_by_pos
        try:
            _, _, rot_rad_map, rot_azi_map = compute_total_rot_polar(mm_to_pos, mm_config_map, alignment_by_pos, gravity_by_pos, thermal_by_pos)
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
            # 1) check the vignetting sheet itself (common exporter puts
            # selected energy into cell C2 or nearby). If present, prefer
            # this value and do not override it from A_eff.
            try:
                # Read cell C2 explicitly from the workbook to avoid DataFrame
                # header/offset issues that can make `iat[1,2]` point to the
                # wrong cell. Use openpyxl to get the literal cell value.
                from openpyxl import load_workbook
                wb_tmp = load_workbook(path, data_only=True)
                if 'Vignetting rotazi' in wb_tmp.sheetnames:
                    ws_tmp = wb_tmp['Vignetting rotazi']
                    candidate = ws_tmp.cell(row=2, column=3).value
                else:
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
                sel_energy = None
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
            print('DBG: starting apply vignetting to rows (will report progress every 100 rows)')
            sys.stdout.flush()
        except Exception:
            pass
        try:
            if os.environ.get('VIG_DEBUG'):
                try:
                    print('DBG: azi_mode=', locals().get('azi_mode'), 'rad_mode=', locals().get('rad_mode'), 'sel_energy=', locals().get('sel_energy'))
                    print('DBG: sample ys_by_pos_azi keys:', list(locals().get('ys_by_pos_azi', {}).keys())[:5])
                    print('DBG: sample ys_by_pos_rad keys:', list(locals().get('ys_by_pos_rad', {}).keys())[:5])
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
                        print(f'DBG: vignetting apply at row {idx}')
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

                # radial
                if applied_rad and p in rot_rad_map:
                    try:
                        # Prefer new per-(cfg_row,energy) layout
                        used = False
                        if 'rad_mode' in locals() and rad_mode == 'per_row_energy' and 'pos_to_cfg_row' in locals():
                            cfg_row = pos_to_cfg_row.get(p)
                            # Use helper to find best-matching per-(cfg_row,energy)
                            if cfg_row is not None:
                                series = _find_series(ys_by_pos_rad, cfg_row, locals().get('sel_energy'), locals().get('aeff_col_name'))
                                if series is not None:
                                    xs_use, ys_use = series
                                    factor = float(np.interp(float(rot_rad_map.get(p, 0.0)), xs_use, ys_use))
                                    used = True
                        if not used and 'rad_mode' in locals() and rad_mode == 'per_pos' and p in ys_by_pos_rad:
                            ys_use = ys_by_pos_rad[p]
                            factor = float(np.interp(float(rot_rad_map.get(p, 0.0)), xs_rad, ys_use))
                        elif not used and 'xs_rad' in locals() and xs_rad is not None and ys_rad is not None:
                            factor = float(np.interp(float(rot_rad_map.get(p, 0.0)), xs_rad, ys_rad))
                        else:
                            factor = 1.0
                        # debug: print chosen key/arrays for first few rows
                        try:
                            if os.environ.get('VIG_DEBUG') and idx < 10:
                                try:
                                    sel_repr = None
                                    try:
                                        if 'keyn' in locals() and keyn in ys_by_pos_rad:
                                            sel_repr = keyn
                                        elif 'key' in locals() and key in ys_by_pos_rad:
                                            sel_repr = key
                                        else:
                                            matches_local = [k for k in ys_by_pos_rad.keys() if k[0] == cfg_row] if 'cfg_row' in locals() else []
                                            sel_repr = matches_local[0] if matches_local else None
                                    except Exception:
                                        sel_repr = None
                                    rot_val = float(rot_rad_map.get(p, 0.0)) if rot_rad_map.get(p) is not None else 0.0
                                    print(f"DBG: RAD idx={idx} pos={p} cfg_row={locals().get('cfg_row')} sel_key={sel_repr} rot_rad={rot_val} factor={factor}")
                                    if sel_repr is not None and sel_repr in ys_by_pos_rad:
                                        xs_print, ys_print = ys_by_pos_rad[sel_repr]
                                        print('DBG: RAD xs[:5]=', xs_print[:5], 'ys[:5]=', ys_print[:5])
                                        try:
                                            print('DBG: RAD interp call ->', rot_val, xs_print.tolist(), ys_print.tolist())
                                            interp_res = np.interp(abs(rot_val), xs_print, ys_print)
                                            print('DBG: RAD interp result=', interp_res)
                                            print('DBG: RAD exact matches indices for rot:', np.where(xs_print == rot_val)[0].tolist())
                                        except Exception as _e:
                                            print('DBG: RAD extra debug failed', _e)
                                    sys.stdout.flush()
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        # record radial factor for this position (always)
                        if 'vig_vals_rad' not in locals():
                            vig_vals_rad = {}
                        if 'vig_source_rad' not in locals():
                            vig_source_rad = {}
                        vig_vals_rad[p] = float(factor)
                        vig_source_rad[p] = ('per_row' if 'rad_mode' in locals() and rad_mode.startswith('per') else 'global')
                    except Exception:
                        factor = 1.0
                        if 'vig_vals_rad' not in locals():
                            vig_vals_rad = {}
                        if 'vig_source_rad' not in locals():
                            vig_source_rad = {}
                        vig_vals_rad[p] = float(factor)
                        vig_source_rad[p] = ('per_row' if 'rad_mode' in locals() and rad_mode.startswith('per') else 'global')

                # azimuthal
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
                            factor = float(np.interp(abs(float(rot_azi_map.get(p, 0.0))), xs_azi, ys_use))
                        elif not used and 'xs_azi' in locals() and xs_azi is not None and ys_azi is not None:
                            factor = float(np.interp(abs(float(rot_azi_map.get(p, 0.0))), xs_azi, ys_azi))
                        else:
                            factor = 1.0
                        # debug: print chosen key/arrays for first few rows (azimuthal)
                        try:
                            if os.environ.get('VIG_DEBUG') and idx < 10:
                                try:
                                    sel_repr = None
                                    try:
                                        if 'keyn' in locals() and keyn in ys_by_pos_azi:
                                            sel_repr = keyn
                                        elif 'key' in locals() and key in ys_by_pos_azi:
                                            sel_repr = key
                                        else:
                                            matches_local = [k for k in ys_by_pos_azi.keys() if k[0] == cfg_row] if 'cfg_row' in locals() else []
                                            sel_repr = matches_local[0] if matches_local else None
                                    except Exception:
                                        sel_repr = None
                                    rot_val = float(rot_azi_map.get(p, 0.0)) if rot_azi_map.get(p) is not None else 0.0
                                    print(f"DBG: AZI idx={idx} pos={p} cfg_row={locals().get('cfg_row')} sel_key={sel_repr} rot_azi={rot_val} factor={factor}")
                                    if sel_repr is not None and sel_repr in ys_by_pos_azi:
                                        xs_print, ys_print = ys_by_pos_azi[sel_repr]
                                        print('DBG: AZI xs[:5]=', xs_print[:5], 'ys[:5]=', ys_print[:5])
                                        try:
                                            print('DBG: AZI interp call ->', rot_val, xs_print.tolist(), ys_print.tolist())
                                            interp_res = np.interp(abs(rot_val), xs_print, ys_print)
                                            print('DBG: AZI interp result=', interp_res)
                                            print('DBG: AZI exact matches indices for rot:', np.where(xs_print == rot_val)[0].tolist())
                                        except Exception as _e:
                                            print('DBG: AZI extra debug failed', _e)
                                    sys.stdout.flush()
                                except Exception:
                                    pass
                        except Exception:
                            pass
                    except Exception:
                        factor = 1.0
                    # record azimuthal factor for this position (do not change df['weight'])
                    if 'vig_vals_azi' not in locals():
                        vig_vals_azi = {}
                    if 'vig_source_azi' not in locals():
                        vig_source_azi = {}
                    vig_vals_azi[p] = float(factor)
                    vig_source_azi[p] = ('per_row' if 'azi_mode' in locals() and azi_mode.startswith('per') else 'global')

        df.attrs['vignetting_rotazi_applied'] = bool(applied_azi)
        df.attrs['vignetting_rotrad_applied'] = bool(applied_rad)
        # Post-pass: recompute per-position vignette values from the
        # populated per-(cfg_row,energy) tables to avoid mismatches that
        # can occur during the row-wise application loop. This ensures
        # `vig_vals_azi` / `vig_vals_rad` reflect the intended
        # interpolation from the vignetting tables.
        try:
            if os.environ.get('VIG_DEBUG'):
                print('DBG: performing post-pass recompute of vig_vals from per-row tables')
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
                        vig_source_azi[pos] = ('per_row' if 'ys_by_pos_azi' in locals() and any(k[0] == cfg for k in ys_by_pos_azi.keys()) else 'global')
                    except Exception:
                        vig_vals_azi[pos] = 1.0
                        vig_source_azi[pos] = 'none'

                    # radial
                    try:
                        applied_val = None
                        if cfg is not None and 'ys_by_pos_rad' in locals() and ys_by_pos_rad:
                            series = _find_series(ys_by_pos_rad, cfg, locals().get('sel_energy'), locals().get('aeff_col_name'))
                            if series is not None:
                                xsr, ysr = series
                                applied_val = float(np.interp(abs(float(rot_rad_map.get(pos, 0.0))), xsr, ysr))
                        if applied_val is None:
                            if 'xs_rad' in locals() and xs_rad is not None and ys_rad is not None:
                                applied_val = float(np.interp(abs(float(rot_rad_map.get(pos, 0.0))), xs_rad, ys_rad))
                            else:
                                applied_val = 1.0
                        vig_vals_rad[pos] = float(applied_val)
                        vig_source_rad[pos] = ('per_row' if 'ys_by_pos_rad' in locals() and any(k[0] == cfg for k in ys_by_pos_rad.keys()) else 'global')
                    except Exception:
                        vig_vals_rad[pos] = 1.0
                        vig_source_rad[pos] = 'none'
        except Exception:
            pass
        try:
            # Compute adjusted A_eff per-MM from aeff_base multiplied by
            # per-position vignetting factors (rotazi * rotrad). Do NOT alter
            # the original A_eff column B in the workbook.
            if 'aeff_base' in df.columns:
                base = df['aeff_base'].astype(float).replace({0.0: np.nan})
                # Build per-position combined factor lookup (default 1.0)
                combined = []
                for mm_idx, row2 in df.iterrows():
                    mmnum = int(row2['MM #']) if not pd.isna(row2['MM #']) else None
                    pos = mm_to_pos.get(int(mmnum)) if mmnum is not None else None
                    f_azi = 1.0
                    f_rad = 1.0
                    if pos is not None:
                        if 'vig_vals_azi' in locals() and pos in vig_vals_azi:
                            f_azi = float(vig_vals_azi.get(pos, 1.0))
                        if 'vig_vals_rad' in locals() and pos in vig_vals_rad:
                            f_rad = float(vig_vals_rad.get(pos, 1.0))
                    combined.append(f_azi * f_rad)
                # Record source info per-MM for debugging
                src_rad = []
                src_azi = []
                for mm_idx, row2 in df.iterrows():
                    mmnum = int(row2['MM #']) if not pd.isna(row2['MM #']) else None
                    pos = mm_to_pos.get(int(mmnum)) if mmnum is not None else None
                    src_rad.append(vig_source_rad.get(pos, 'none') if 'vig_source_rad' in locals() else 'none')
                    src_azi.append(vig_source_azi.get(pos, 'none') if 'vig_source_azi' in locals() else 'none')
                df['aeff_vig_source_rotrad'] = src_rad
                df['aeff_vig_source_rotazi'] = src_azi
                df['aeff_vig_factor'] = np.array(combined, dtype=float)
                df['aeff_adjusted'] = base * df['aeff_vig_factor']
                # Make adjusted A_eff the active per-MM weight for aggregation
                try:
                    df['weight'] = df['aeff_adjusted'].astype(float)
                except Exception:
                    pass
        except Exception:
            pass

        # Attempt to write per-position vignette factors into the workbook's
        # Vignetting rotazi/rotrad sheets in column B for visibility. This
        # implementation uses openpyxl to update only the specified cells
        # (column B and C1 on vignette sheets; columns B and C on A_eff)
        # and saves atomically to avoid corrupting the original workbook.
        try:
            from openpyxl import load_workbook
            print('DBG: using openpyxl in-place vignette writer')
            sys.stdout.flush()
            wb = load_workbook(path)

            # VIGNETTE SHEETS: write col B and C1 only
            for sname, vig_map in (
                ('Vignetting rotazi', vig_vals_azi if 'vig_vals_azi' in locals() else {}),
                ('Vignetting rotrad', vig_vals_rad if 'vig_vals_rad' in locals() else {}),
            ):
                # Debug print of sample vig_map contents when requested
                try:
                    if os.environ.get('VIG_DEBUG'):
                        sample_items = list((vig_vals_azi if sname.endswith('rotazi') else vig_vals_rad).items())[:10]
                        print(f"DBG: sample vig_map for {sname}: {sample_items}")
                        sys.stdout.flush()
                except Exception:
                    pass

                # Extra debug: show a few per-position arrays and rot offsets for pos 1
                try:
                    if os.environ.get('VIG_DEBUG'):
                        p0 = 1
                        print('DBG: rot_rad_map[1]=', rot_rad_map.get(1))
                        print('DBG: rot_azi_map[1]=', rot_azi_map.get(1))
                        if 'pos_to_cfg_row' in locals():
                            cfg = pos_to_cfg_row.get(p0)
                            print('DBG: pos_to_cfg_row[1]=', cfg)
                            if 'ys_by_pos_azi' in locals():
                                matches = [k for k in ys_by_pos_azi.keys() if k[0] == cfg]
                                print('DBG: azi matches for cfg:', matches[:5])
                                if matches:
                                    k = matches[0]
                                    xsu, ysu = ys_by_pos_azi[k]
                                    print('DBG: sample azi xs[:5]=', xsu[:5], 'ys[:5]=', ysu[:5])
                            if 'ys_by_pos_rad' in locals():
                                matchesr = [k for k in ys_by_pos_rad.keys() if k[0] == cfg]
                                print('DBG: rad matches for cfg:', matchesr[:5])
                                if matchesr:
                                    kr = matchesr[0]
                                    xsr, ysr = ys_by_pos_rad[kr]
                                    print('DBG: sample rad xs[:5]=', xsr[:5], 'ys[:5]=', ysr[:5])
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
                        print(f"DBG: writing vignette values progress: {i}/{total} for sheet '{sname}'")
                        sys.stdout.flush()
                    pos_int = int(pos_k)
                    # prefer already-computed vig_map value
                    val = None
                    try:
                        if pos_int in vig_map:
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

                print(f"DBG: wrote {written} values into sheet '{sname}'")

                # Do not write to cell C1 here; GUI is responsible for any
                # selected-energy markers (e.g. cell C2). Leave C1/C2 untouched.

            # A_eff sheet: update columns B and C only
            if 'A_eff' in wb.sheetnames and ('aeff_adjusted' in df.columns or 'aeff_base' in df.columns):
                ws_a = wb['A_eff']
                max_r_a = ws_a.max_row or 0

                mm_series = pd.to_numeric(df.get('MM #', pd.Series([], dtype=float)), errors='coerce')
                base_series = pd.to_numeric(df.get('aeff_base', pd.Series([], dtype=float)), errors='coerce')
                adj_series = pd.to_numeric(df.get('aeff_adjusted', pd.Series([], dtype=float)), errors='coerce')

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
                    wrote = False
                    if mmv in mm_to_base:
                        try:
                            ws_a.cell(row=r, column=2, value=float(mm_to_base[mmv]))
                            wrote = True
                        except Exception:
                            pass
                    if mmv in mm_to_adj:
                        try:
                            ws_a.cell(row=r, column=3, value=float(mm_to_adj[mmv]))
                            wrote = True
                        except Exception:
                            pass
                    if wrote:
                        written_a += 1

                if written_a > 0:
                    # Save atomically to avoid corruption
                    tmpf = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                    tmpf.close()
                    wb.save(tmpf.name)
                    os.replace(tmpf.name, path)
                    print(f"DBG: wrote {written_a} rows into 'A_eff' columns B/C")
                    sys.stdout.flush()

            try:
                wb.close()
            except Exception:
                pass
            print('DBG: openpyxl in-place vignette write completed')
            sys.stdout.flush()
        except Exception:
            # If anything goes wrong, do not raise — vignetting writes are non-fatal.
            print('DBG: in-place vignette writer failed; skipping writes')
            sys.stdout.flush()
        # Debug summary: print selected vignette source mapping for first positions
        try:
            debug_env = os.environ.get('VIG_DEBUG', None)
            if debug_env:
                sample = sorted(list(mm_to_pos.values()))[:16]
                print('VIGNETTE DEBUG: per-position sources (pos: rad_source, azi_source)')
                for pos in sample:
                    rsrc = vig_source_rad.get(pos, 'none') if 'vig_source_rad' in locals() else 'none'
                    asrc = vig_source_azi.get(pos, 'none') if 'vig_source_azi' in locals() else 'none'
                    print(f' pos {pos}: {rsrc}, {asrc}')
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
        weight_arr_for_center = df['aeff_adjusted'].to_numpy(dtype=float, copy=False)
    elif 'weight' in df.columns:
        weight_arr_for_center = df['weight'].to_numpy(dtype=float, copy=False)
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
    
    # --- Fast grid summation helpers (threaded) ---
    mux_arr = df['mux'].to_numpy(dtype=float, copy=False)
    muy_arr = df['muy'].to_numpy(dtype=float, copy=False)
    sigx_arr = df['sigmax'].to_numpy(dtype=float, copy=False)
    sigy_arr = df['sigmay'].to_numpy(dtype=float, copy=False)
    theta_arr = df['theta_degrees'].to_numpy(dtype=float, copy=False)
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
                Zc += pseudo_voigt_2d_rotated(
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
            elif dlow == 'gaussian':
                Zc += gaussian_2d_rotated(
                    Xg, Yg,
                    mux=mux_arr[i], muy=muy_arr[i],
                    sigmax=sigx_arr[i], sigmay=sigy_arr[i],
                    theta=theta_arr[i],
                    amplitude=weight_arr[i],
                    normalize=normalize_flag,
                    degrees=True,
                )
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
            try:
                frac = cumulative / total_energy if total_energy > 0 else cumulative
                import numpy as _np
                print(f"[plot_sum.debug] radial_profile: cx={cx:.6e}, cy={cy:.6e}, r_max={r_max:.6e}, max_sigma={max_sigma:.6e}, max_center_dist={max_center_dist:.6e}, n_r={n_r}, n_theta={n_theta}, total_energy={total_energy:.6e}, frac0={float(frac[0]) if frac.size>0 else None}, frac_end={float(frac[-1]) if frac.size>0 else None}, frac_nans={int(_np.isnan(frac).sum())}")
            except Exception:
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

    # Also compute 50% from origin for reference
    r_profile_00, cumulative_00, total_00 = radial_profile(0.0, 0.0, n_r=n_r_final, n_theta=n_theta_final, r_margin_factor=final_r_margin)
    frac_00 = cumulative_00 / total_00 if total_00 > 0 else cumulative_00
    radius_50_00 = _radius_for_fraction(frac_00, r_profile_00, target=0.5)
    # 90% at origin for reference
    radius_90_00 = _radius_for_fraction(frac_00, r_profile_00, target=0.9)
    
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
                try:
                    frac = cumulative / total_energy if total_energy > 0 else cumulative
                    import numpy as _np
                    print(f"[plot_sum.debug] radial_profile_opt: cx={cx:.6e}, cy={cy:.6e}, r_max={r_max:.6e}, max_sigma={max_sigma:.6e}, max_center_dist={max_center_dist:.6e}, n_r={n_r}, n_theta={n_theta}, total_energy={total_energy:.6e}, frac0={float(frac[0]) if frac.size>0 else None}, frac_end={float(frac[-1]) if frac.size>0 else None}, frac_nans={int(_np.isnan(frac).sum())}")
                except Exception:
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
    # Use a reasonable figure size that will fit most screens
    fig = plt.figure(figsize=(16, 9))
    
    # Use GridSpec with 20 columns for finer control: 65% left (13 cols) + 35% right (7 cols)
    gs = gridspec.GridSpec(1, 20)
    
    # First subplot: weighted sum of Gaussians (65% width)
    ax1 = plt.subplot(gs[0, :13])
    # Convert to microns for display (1 m = 1e6 µm)
    im = plt.imshow(Z, extent=[x.min()*1e6, x.max()*1e6, y.min()*1e6, y.max()*1e6], origin='lower', cmap='viridis', aspect='equal')
    cbar = plt.colorbar(im, label='counts', pad=0.12)
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
    
    # Mark the best focus with a green cross and coordinates
    plt.plot(center_x*1e6, center_y*1e6, 'gx', markersize=10, label='best focus')
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
    plt.xlim((xlim_min - margin_x)*1e6, (xlim_max + margin_x)*1e6)
    plt.ylim((ylim_min - margin_y)*1e6, (ylim_max + margin_y)*1e6)
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
        return best

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

    hew_opt_x_arcsec, hew_opt_y_arcsec = (None, None)
    if df_optimized is not None:
        # Reuse the same plotting grid for the overlay dataset.
        Z_opt = _sum_on_grid_opt(X, Y, normalize)
        hew_opt_x_arcsec, hew_opt_y_arcsec = _compute_hew_xy_arcsec_from_grid_marginals(x, y, Z_opt)
    hew_best_arcsec = 2 * radius_50 * m_to_arcsec if radius_50 is not None else None
    hew_origin_arcsec = 2 * radius_50_00 * m_to_arcsec if radius_50_00 is not None else None
    eef90_arcsec = 2 * radius_90 * m_to_arcsec if radius_90 is not None else None
    eef90_origin_arcsec = 2 * radius_90_00 * m_to_arcsec if radius_90_00 is not None else None

    # If caller only requests metrics, return them now without any plotting side-effects.
    if return_metrics_only:
        return {
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
    else:
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
        # Add a dashed circle for 90% encircled energy in red
        label_90 = f'EEF 90% centered on best focus ({eef90_arcsec:.4f}" diameter)' if eef90_arcsec is not None else 'EEF 90% centered on best focus'
        circle_90 = plt.Circle((center_x*1e6, center_y*1e6), radius_90*1e6, fill=False, color='red', linestyle='--', linewidth=2, label=label_90)
        plt.gca().add_patch(circle_90)
    if radius_50_00 is not None:
        # Add a dashed circle for 50% encircled energy from (0,0) in blue
        label_50_00 = f'HEW centered on (0,0) ({hew_origin_arcsec:.4f}" diameter)' if hew_origin_arcsec is not None else 'HEW centered on (0,0)'
        circle_50_00 = plt.Circle((0, 0), radius_50_00*1e6, fill=False, color='blue', linestyle='--', linewidth=2, label=label_50_00)
        plt.gca().add_patch(circle_50_00)
    if radius_50 is not None:
        # Add a dashed circle for 50% encircled energy in green
        label_50 = f'HEW centered on best focus ({hew_best_arcsec:.4f}" diameter)' if hew_best_arcsec is not None else 'HEW centered on best focus'
        circle_50 = plt.Circle((center_x*1e6, center_y*1e6), radius_50*1e6, fill=False, color='green', linestyle='--', linewidth=2, label=label_50)
        plt.gca().add_patch(circle_50)
    
    # Add optimized circles if provided
    if df_optimized is not None and opt_radius_50 is not None:
        opt_hew_arcsec = 2 * opt_radius_50 * m_to_arcsec
        opt_eef90_arcsec = 2 * opt_radius_90 * m_to_arcsec if opt_radius_90 is not None else None
        # Mark optimized best focus
        plt.plot(opt_center_x*1e6, opt_center_y*1e6, 'mx', markersize=10, label='optimized best focus')
        # HEW circle (magenta dotted)
        label_opt_50 = f'HEW optimized ({opt_hew_arcsec:.4f}" diameter)'
        circle_opt_50 = plt.Circle((opt_center_x*1e6, opt_center_y*1e6), opt_radius_50*1e6, fill=False, color='magenta', linestyle=':', linewidth=2, label=label_opt_50)
        plt.gca().add_patch(circle_opt_50)
        # EEF 90% circle (orange dotted)
        if opt_radius_90 is not None:
            label_opt_90 = f'EEF 90% optimized ({opt_eef90_arcsec:.4f}" diameter)'
            circle_opt_90 = plt.Circle((opt_center_x*1e6, opt_center_y*1e6), opt_radius_90*1e6, fill=False, color='orange', linestyle=':', linewidth=2, label=label_opt_90)
            plt.gca().add_patch(circle_opt_90)
    
    # Reorder legend items on the left subplot:
    # 1) focus points, 2) HEW circles, 3) EEF 90% circles
    handles, labels = ax1.get_legend_handles_labels()

    focus_order = {
        'best focus': 0,
        '(0,0)': 1,
        'optimized best focus': 2,
    }

    def _legend_sort_key(label: str) -> tuple[int, int, str]:
        if label in focus_order:
            return (0, focus_order[label], label)
        if label.startswith('HEW'):
            if 'best focus' in label:
                return (1, 0, label)
            if '(0,0)' in label:
                return (1, 1, label)
            if 'optimized' in label:
                return (1, 2, label)
            return (1, 99, label)
        if label.startswith('EEF 90%'):
            if 'best focus' in label:
                return (2, 0, label)
            if 'optimized' in label:
                return (2, 1, label)
            return (2, 99, label)
        return (3, 99, label)

    order = sorted(range(len(labels)), key=lambda i: _legend_sort_key(labels[i]))
    handles_sorted = [handles[i] for i in order]
    labels_sorted = [labels[i] for i in order]
    ax1.legend(handles_sorted, labels_sorted, loc='upper right')
    
    # Second subplot: encircled energy function (35% width)
    ax2 = plt.subplot(gs[0, 13:])
    # Convert diameter from meters to arcsec (1 m = 54000/π arcsec)
    profile_pct = frac_profile * 100 if 'frac_profile' in locals() else []
    profile_diam = 2 * r_profile * m_to_arcsec if 'r_profile' in locals() else []
    profile_pct_00 = frac_00 * 100 if 'frac_00' in locals() else []
    profile_diam_00 = 2 * r_profile_00 * m_to_arcsec if 'r_profile_00' in locals() else []
    label_best = 'Centered on best focus'
    if (hew_best_x_arcsec is not None) and (hew_best_y_arcsec is not None):
        label_best = f'Centered on best focus (HEW_x={hew_best_x_arcsec:.4f}", HEW_y={hew_best_y_arcsec:.4f}")'
    label_00 = 'Centered on (0,0)'
    if (hew_00_x_arcsec is not None) and (hew_00_y_arcsec is not None):
        label_00 = f'Centered on (0,0) (HEW_x={hew_00_x_arcsec:.4f}", HEW_y={hew_00_y_arcsec:.4f}")'
    plt.plot(profile_pct, profile_diam, label=label_best, color='green')
    plt.plot(profile_pct_00, profile_diam_00, label=label_00, color='blue')
    
    # Add optimized curve if provided
    if df_optimized is not None and opt_frac_profile is not None:
        opt_profile_pct = opt_frac_profile * 100
        opt_profile_diam = 2 * opt_r_profile * m_to_arcsec
        label_opt = 'Optimized best focus'
        if (hew_opt_x_arcsec is not None) and (hew_opt_y_arcsec is not None):
            label_opt = f'Optimized best focus (HEW_x={hew_opt_x_arcsec:.4f}", HEW_y={hew_opt_y_arcsec:.4f}")'
        plt.plot(opt_profile_pct, opt_profile_diam, label=label_opt, linestyle=':', linewidth=2.5, color='magenta')
    
    ax2.set_xlabel('Percentage (%)')
    ax2.set_ylabel('Diameter [arcsec]')
    ax2.legend()
    # Increase tick density by reducing major tick spacing by half
    from matplotlib.ticker import MultipleLocator, AutoMinorLocator
    ax2.xaxis.set_major_locator(MultipleLocator(10))  # Major ticks every 10%
    ax2.xaxis.set_minor_locator(AutoMinorLocator(5))  # Minor ticks
    # For y-axis, use auto locator with more ticks
    ax2.yaxis.set_major_locator(plt.MaxNLocator(nbins=20))  # More bins for denser ticks
    ax2.yaxis.set_minor_locator(AutoMinorLocator(5))
    # Mark the 50% encircled energy (Half Energy Width - HEW) in green for best focus
    plt.axhline(y=hew_best_arcsec, linestyle='--', color='green')  # Horizontal line at HEW diameter
    plt.axvline(x=50, linestyle='--', color='black')  # Vertical line at 50%
    if hew_best_arcsec is not None:
        plt.text(0, hew_best_arcsec, f'HEW best focus: {hew_best_arcsec:.4f}"', ha='left', va='top', fontsize=10, color='green')  # Label HEW with value
    # Mark the 50% from (0,0) in blue
    plt.axhline(y=hew_origin_arcsec, linestyle='--', color='blue')  # Horizontal line at HEW(0,0) diameter
    if hew_origin_arcsec is not None:
        plt.text(100, hew_origin_arcsec, f'HEW (0,0): {hew_origin_arcsec:.4f}"', ha='center', va='top', fontsize=10, color='blue')  # Label HEW (0,0) with value
    # Mark the 90% encircled energy in red
    if radius_90 is not None:
        plt.axhline(y=eef90_arcsec, linestyle='--', color='red')  # Horizontal line at EEF90 diameter
        plt.axvline(x=90, linestyle='--', color='black')  # Vertical line at 90%
        plt.text(0, eef90_arcsec, f'EEF 90% best focus: {eef90_arcsec:.4f}"', ha='left', va='top', fontsize=10, color='red')  # Label 90% with value
    
    # Add optimized reference lines if provided
    if df_optimized is not None and opt_radius_50 is not None:
        opt_hew_arcsec = 2 * opt_radius_50 * m_to_arcsec
        plt.axhline(y=opt_hew_arcsec, linestyle=':', color='magenta', linewidth=1.5)
        plt.text(0, opt_hew_arcsec, f'HEW optimized: {opt_hew_arcsec:.4f}"', ha='left', va='bottom', fontsize=10, color='magenta')
        if opt_radius_90 is not None:
            opt_eef90_arcsec = 2 * opt_radius_90 * m_to_arcsec
            plt.axhline(y=opt_eef90_arcsec, linestyle=':', color='orange', linewidth=1.5)
            plt.text(0, opt_eef90_arcsec, f'EEF 90% optimized: {opt_eef90_arcsec:.4f}"', ha='left', va='bottom', fontsize=10, color='orange')
    
    # Add titles at the same y-coordinate
    fig.suptitle('')  # Clear any figure title
    ax1.set_title(f'E2E PSF{title_suffix}', fontweight='bold', fontsize=24, y=1.08)
    ax2.set_title(f'Encircled energy function{title_suffix}', fontweight='bold', fontsize=24, y=1.08)
    
    plt.tight_layout()  # Adjust layout
    
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
            pct_best = profile_pct if 'profile_pct' in locals() else (frac_profile * 100 if 'frac_profile' in locals() else [])
            diam_best = profile_diam if 'profile_diam' in locals() else (2 * r_profile * m_to_arcsec if 'r_profile' in locals() else [])
            pct_orig = profile_pct_00 if 'profile_pct_00' in locals() else (frac_00 * 100 if 'frac_00' in locals() else [])
        except Exception:
            pct_best, diam_best, pct_orig, diam_orig = [], [], [], []

        # Optimized arrays optional
        opt_pct = None
        opt_diam = None
        if 'opt_frac_profile' in locals() and opt_frac_profile is not None:
            try:
                opt_pct = opt_frac_profile * 100
                opt_diam = 2 * opt_r_profile * m_to_arcsec
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
        
        menu_text = "┌─────────────────────────────┐\n"
        menu_text += "│  1. Export PSF Plot         │\n"
        menu_text += "│  2. Export EEF Plot         │\n"
        menu_text += "│  3. Export FITS             │\n"
        menu_text += "│  4. Export EEF CSV         │\n"
        menu_text += "│  5. Cancel                  │\n"
        menu_text += "└─────────────────────────────┘"
        
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
                    
                    # Menu structure for 5 options (top-to-bottom):
                    # Top border, Opt1, Opt2, Opt3, Opt4, Opt5, Bottom border
                    # Ranges chosen empirically to match text layout
                    if 0.80 < relative_y <= 0.95:  # Option 1 (Export PSF)
                        hide_context_menu()
                        export_psf_plot()
                    elif 0.62 < relative_y <= 0.80:  # Option 2 (Export EEF)
                        hide_context_menu()
                        export_eef_plot()
                    elif 0.44 < relative_y <= 0.62:  # Option 3 (Export FITS)
                        hide_context_menu()
                        export_fits()
                    elif 0.26 < relative_y <= 0.44:  # Option 4 (Export EEF CSV)
                        hide_context_menu()
                        export_eef_csv()
                    elif 0.08 < relative_y <= 0.26:  # Option 5 (Cancel)
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
                export_eef_csv()
            elif event.key in ['5', 'escape']:
                hide_context_menu()
        else:
            if event.key in ['p', '1']:
                export_psf_plot()
            elif event.key in ['e', '2']:
                export_eef_plot()
            elif event.key in ['f', '3']:
                export_fits()
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
            try:
                plt.show()
            except Exception:
                print("Interactive display failed — no GUI available. To save the plot use --output <file.png> or run in a GUI session.")


def compute_hew_eef_metrics(file: str = 'Distributions/Test_Distribution.xlsx', sheet: str = 'MM_PSF', normalize: bool = True, fast: bool = True, df_optimized: pd.DataFrame = None) -> dict:
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
    parser.add_argument('-f','--file', default='Distributions/Test_Distribution.xlsx', help='Excel file path')
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
        except Exception:
            pass

    # Load data from Excel
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
