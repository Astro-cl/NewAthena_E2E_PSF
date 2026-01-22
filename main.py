
import argparse  # For parsing command-line arguments
import numpy as np  # For numerical operations and arrays
import pandas as pd  # For data manipulation and Excel reading
import matplotlib.pyplot as plt  # For plotting
import matplotlib.gridspec as gridspec
import os
from concurrent.futures import ThreadPoolExecutor
from distributions_rotated import (
    gaussian_2d_rotated,
    pseudo_voigt_2d_rotated,
    load_psf_matrix_excel,
    eval_psf_matrix_rotated,
)  # Custom functions for 2D rotated distributions
import json
import sys


def load_aeff_weight_map(path: str, sheet: str = 'A_eff') -> dict[int, float]:
    """Load weight mapping from A_eff sheet: {MM # -> weight}.

    Contract (strict):
    - Weights MUST come from column B of the 'A_eff' sheet.
    - MM identifiers MUST come from column A of the same sheet.
    - If the sheet is missing, or any MM has a missing/non-numeric weight in column B,
      this function raises ValueError.
    """
    try:
        raw = pd.read_excel(path, sheet_name=sheet, engine='openpyxl', header=None)
    except Exception as e:
        raise ValueError(f"Missing or unreadable '{sheet}' sheet in workbook: {e}")

    if raw.shape[1] < 2:
        raise ValueError(f"'{sheet}' sheet must have at least 2 columns (A=MM #, B=weight).")

    # Find header row containing 'MM #' in column A, else assume first row is header.
    header_row = 0
    scan_rows = min(20, raw.shape[0])
    for r in range(scan_rows):
        v = raw.iloc[r, 0]
        if isinstance(v, str) and v.strip().lower().replace(' ', '') in {'mm#', 'mm'}:
            header_row = r
            break

    data = raw.iloc[header_row + 1 :, :2].copy()
    data.columns = ['MM #', 'weight']

    mm = pd.to_numeric(data['MM #'], errors='coerce')
    wt = pd.to_numeric(data['weight'], errors='coerce')

    # Rows with no MM are ignored; rows with an MM must have a valid numeric weight.
    has_mm = mm.notna()
    invalid_weight = has_mm & wt.isna()
    if invalid_weight.any():
        bad_rows = data.loc[invalid_weight, ['MM #', 'weight']].head(10).copy()
        # Preserve the original raw row index to report Excel row numbers.
        # raw/data are 0-indexed; Excel rows are 1-indexed.
        excel_rows = [int(i) + 1 for i in bad_rows.index.tolist()]
        bad_rows.insert(0, 'Excel row', excel_rows)
        bad_cells = [f"{sheet}!B{r}" for r in excel_rows]
        raise ValueError(
            "Invalid A_eff weights in column B for some rows. "
            "Column B must be numeric for every MM. "
            f"Bad cells (examples): {', '.join(bad_cells)}\n"
            "Example bad rows:\n"
            + bad_rows.to_string(index=False)
        )

    mm = mm[has_mm].astype(int)
    wt = wt[has_mm].astype(float)
    if mm.empty:
        raise ValueError(f"'{sheet}' sheet has no MM entries in column A.")

    # Detect duplicates / conflicts
    tmp = pd.DataFrame({'MM #': mm.to_numpy(), 'weight': wt.to_numpy()})
    if tmp.duplicated(subset=['MM #']).any():
        # Allow duplicates only if weights are identical
        grp = tmp.groupby('MM #')['weight'].nunique(dropna=True)
        conflict = grp[grp > 1]
        if not conflict.empty:
            raise ValueError(f"Conflicting weights found for MM(s): {conflict.index.tolist()[:20]}")
        tmp = tmp.drop_duplicates(subset=['MM #'], keep='first')

    return dict(zip(tmp['MM #'].astype(int), tmp['weight'].astype(float)))


def load_gaussians_from_excel(path: str, sheet: str | None = None) -> pd.DataFrame:
    """Load gaussian parameters from Excel.

    Expected columns: m_rad [arcsec], m_azi [arcsec], sigma_rad [arcsec], sigma_azi [arcsec]
    Values are converted from arcsec to meters using: 1 arcsec = 12*π/180/3600 m
    Converts to mux, muy using rotation matrix:
    - mux = cos(theta)*m_rad - sin(theta)*m_azi
    - muy = sin(theta)*m_rad + cos(theta)*m_azi
    theta_degrees is calculated from MM configuration: theta = arcsin(x_MM / r_MM)
    weight column is optional and will be overridden by A_eff sheet if present
    """
    kwargs = {"engine": "openpyxl"}  # Use openpyxl engine for Excel files
    if sheet:
        kwargs["sheet_name"] = sheet  # Specify sheet if provided
    df = pd.read_excel(path, **kwargs)  # Read the Excel file into a DataFrame
    required = ["m_rad [arcsec]","m_azi [arcsec]","sigma_rad [arcsec]","sigma_azi [arcsec]"]  # Required columns
    if not all(c in df.columns for c in required):  # Check if all required columns are present
        raise ValueError(f"Excel must contain columns: {required}")  # Raise error if not
    
    # Convert from arcsec to meters: 1 arcsec = 12*π/180/3600 m
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
        aeff_map = load_aeff_weight_map(path)

        mm_as_int = pd.to_numeric(df['MM #'], errors='coerce')
        if mm_as_int.isna().any():
            bad = df.loc[mm_as_int.isna(), 'MM #'].head(10).tolist()
            raise ValueError(f"Invalid 'MM #' values in PSF sheet: {bad}")
        mm_as_int = mm_as_int.astype(int)

        df['weight'] = mm_as_int.map(aeff_map)
        missing_mask = df['weight'].isna()
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
                    alignment_by_pos[pos] = {
                        'd_align_rad': float(row.get('d_align_rad [µm]', 0)) * 1e-6,
                        'd_align_azi': float(row.get('d_align_azi [µm]', 0)) * 1e-6,
                        'd_align_z': float(row.get('d_align_z [µm]', 0)) * 1e-6,
                        'd_align_rotz': float(row.get('d_align_rotz [arcsec]', 0)),
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
                for c in ['d_grav_x [µm]', 'd_grav_y [µm]', 'd_grav_z [µm]', 'd_grav_rotz [arcsec]']:
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
                for c in ['d_therm_x [µm]', 'd_therm_y [µm]', 'd_therm_z [µm]', 'd_therm_rotz [arcsec]']:
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
                    }
        except Exception:
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
        if mm_num not in mm_config_map:
            return row['mux'], row['muy']  # Use default if no config
        
        config = mm_config_map[mm_num]
        x_mm = config.get('x_MM', 0)
        y_mm = config.get('y_MM', 0)
        r_mm = config.get('r_MM', 1)  # Avoid division by zero
        
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
    if 'weight' in df.columns:
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
        # Coarse: keep radial emphasis but shrink azimuthal samples to meet runtime.
        n_r_final = 10000
        n_theta_final = 120
        final_r_margin = 22.0
    else:
        # Increase fine-mode sampling for higher accuracy (budget ~15s).
        n_r_final = 12000
        n_theta_final = 1440
        final_r_margin = 22.0

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
    
    # Add secondary axes for arcsec (1 arcsec = 12*π/180/3600 m, so 1 m = 54000/π arcsec)
    m_to_arcsec = 54000 / np.pi  # meters to arcsec
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
    m_to_arcsec = 54000 / np.pi

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
        i = 0
        for j in range(1, prof.size):
            # Increase i while we still meet target
            while i < j and (pref[j] - pref[i]) >= target:
                width = float(axis_vals[j] - axis_vals[i])
                if (best is None) or (width < best):
                    best = width
                i += 1
        return best

    def _compute_hew_xy_arcsec_from_grid_marginals(x_axis: np.ndarray, y_axis: np.ndarray, Zg: np.ndarray) -> tuple[float | None, float | None]:
        """Compute HEW_x and HEW_y (arcsec) from the full summed 2D PSF Zg.

        Uses *marginals* of the aggregated PSF:
        - Px(x) = \int Z(x,y) dy
        - Py(y) = \int Z(x,y) dx

        Then computes the minimal-width interval containing 50% energy for each marginal.
        """
        if Zg.size == 0:
            return (None, None)
        if not np.isfinite(Zg).any():
            return (None, None)

        x_axis = np.asarray(x_axis, dtype=float)
        y_axis = np.asarray(y_axis, dtype=float)
        Zg = np.asarray(Zg, dtype=float)

        # Marginals (integrate the full 2D PSF along the other axis)
        prof_x = np.trapezoid(Zg, y_axis, axis=0)
        prof_y = np.trapezoid(Zg, x_axis, axis=1)

        hew_x_m = _min_interval_width(x_axis, prof_x, frac=0.5)
        hew_y_m = _min_interval_width(y_axis, prof_y, frac=0.5)

        hew_x_arcsec = (hew_x_m * m_to_arcsec) if hew_x_m is not None else None
        hew_y_arcsec = (hew_y_m * m_to_arcsec) if hew_y_m is not None else None
        return hew_x_arcsec, hew_y_arcsec

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
    
    def show_context_menu(x, y):
        """Display context menu at given position"""
        nonlocal menu_annotation, menu_active
        
        if menu_active:
            hide_context_menu()
            return
        
        menu_text = "┌─────────────────────────────┐\n"
        menu_text += "│  1. Export PSF Plot         │\n"
        menu_text += "│  2. Export EEF Plot         │\n"
        menu_text += "│  3. Cancel                  │\n"
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
                    
                    # Menu structure (from top to bottom):
                    # Top border, Option 1, Option 2, Option 3, Bottom border
                    # Inverted because we calculate from bottom: higher y = higher in menu
                    if 0.60 < relative_y < 0.85:  # Option 1 (Export PSF)
                        hide_context_menu()
                        export_psf_plot()
                    elif 0.40 < relative_y < 0.60:  # Option 2 (Export EEF)
                        hide_context_menu()
                        export_eef_plot()
                    elif 0.15 < relative_y < 0.40:  # Option 3 (Cancel)
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
            elif event.key in ['3', 'escape']:
                hide_context_menu()
        else:
            if event.key in ['p', '1']:
                export_psf_plot()
            elif event.key in ['e', '2']:
                export_eef_plot()
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
    
    if output:
        plt.savefig(output, dpi=150)  # Save the combined plot
        print(f"Saved combined plot to {output}")

    if not output:
        plt.show()  # Display the combined plot


def compute_hew_eef_metrics(file: str = 'Distributions/Test_Distribution.xlsx', sheet: str = 'MM_PSF', normalize: bool = True, fast: bool = True, df_optimized: pd.DataFrame = None) -> dict:
    """Convenience wrapper: load workbook and return HEW/EEF metrics (no plotting).

    Returns a dict matching the CLI JSON output.
    """
    df = load_gaussians_from_excel(file, sheet)
    return plot_sum(df, normalize=normalize, fast=fast, df_optimized=df_optimized, return_metrics_only=True)


if __name__ == '__main__':
    # Close all existing figures to ensure clean start
    plt.close('all')
    
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='Plot sum of rotated Gaussians from Excel.')
    parser.add_argument('-f','--file', default='Distributions/Test_Distribution.xlsx', help='Excel file path')
    parser.add_argument('-s','--sheet', default='MM_PSF', help='Sheet name to read (default MM_PSF)')
    parser.add_argument('--normalize', dest='normalize', action='store_true', default=True, help='Normalize each Gaussian to integrate to 1 (default on)')
    parser.add_argument('--no-normalize', dest='normalize', action='store_false', help='Disable normalization')
    parser.add_argument('-o','--output', help='Optional output image file path (PNG)')
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
    args = parser.parse_args()

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
            df_optimized = load_gaussians_from_excel(opt_output, args.sheet)
            plot_title_suffix = " (comparison)"

    # Placement-only mode
    elif placement_strategy is not None:
        from optimize_mm_rows import cross_placement, x_axis_placement, elliptical_placement
        base, ext = os.path.splitext(args.file)
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
            else:
                df_optimized = None
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
