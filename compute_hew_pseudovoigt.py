
import argparse
import pandas as pd
import numpy as np
from optimize_mm_rows import compute_mm_hew_at_origin

def main():
    parser = argparse.ArgumentParser(description="Compute HEW for a 2D pseudo-Voigt distribution.")
    parser.add_argument('--sigma_rad', type=float, required=True, help='Radial sigma (e.g., 3.8)')
    parser.add_argument('--sigma_azi', type=float, required=True, help='Azimuthal sigma (e.g., 1.0)')
    parser.add_argument('--alpha_rad', type=float, required=True, help='Radial alpha (e.g., 0.77)')
    parser.add_argument('--alpha_azi', type=float, required=True, help='Azimuthal alpha (e.g., 0.29)')
    args = parser.parse_args()

    # Convert input sigmas from arcsec to meters (main.py convention)
    arcsec_to_m = 12 * np.pi / 180 / 3600
    sigma_rad_m = args.sigma_rad * arcsec_to_m
    sigma_azi_m = args.sigma_azi * arcsec_to_m
    df = pd.DataFrame({
        "MM #": [1],
        "m_rad": [0.0],
        "m_azi": [0.0],
        "sigma_rad": [sigma_rad_m],
        "sigma_azi": [sigma_azi_m],
        "theta_degrees": [0.0],
        "weight": [1.0],
        "distribution": ["pseudo-voigt"],
        "alpha_rad": [args.alpha_rad],
        "alpha_azi": [args.alpha_azi],
    })
    # Use high-resolution grid for HEW calculation
    from optimize_mm_rows import hew_fast_approximate_center
    # Extract parameters for single MM
    row = df.iloc[0]
    single = pd.DataFrame([
        {
            'MM #': 1,
            'mux': 0.0,
            'muy': 0.0,
            'sigmax': row['sigma_rad'],
            'sigmay': row['sigma_azi'],
            'theta_degrees': row['theta_degrees'],
            'weight': row['weight'],
            'distribution': row['distribution'],
            'alpha_azi': row['alpha_azi'],
            'alpha_rad': row['alpha_rad'],
        }
    ])
    # Set grid resolution
    n_r = 360
    n_theta = 200
    # Build polar grid
    max_sigma = max(single['sigmax'].max(), single['sigmay'].max())
    r_max = max(0.0 + 4.0 * max_sigma, 1e-6)
    theta = np.linspace(0.0, 2.0 * np.pi, n_theta, endpoint=False)
    r = np.linspace(0.0, r_max, n_r)
    dtheta = theta[1] - theta[0]
    dr = r[1] - r[0] if n_r > 1 else r_max
    R, TH = np.meshgrid(r, theta)
    Xp = 0.0 + R * np.cos(TH)
    Yp = 0.0 + R * np.sin(TH)
    Zp = np.zeros_like(Xp)
    for _, row in single.iterrows():
        dist_type = row.get('distribution', 'gaussian')
        if dist_type in ['pseudo-voigt', 'voigt']:
            from distributions_rotated import pseudo_voigt_2d_rotated
            Zp += pseudo_voigt_2d_rotated(
                Xp, Yp,
                muazi=row['mux'], murad=row['muy'],
                sigmaazi=row['sigmax'], sigmarad=row['sigmay'],
                theta=row['theta_degrees'],
                alphaazi=row.get('alpha_azi', 0.5),
                alpharad=row.get('alpha_rad', 0.5),
                amplitude=row['weight'],
                normalize=True,
                degrees=True,
            )
        else:
            from distributions_rotated import gaussian_2d_rotated
            Zp += gaussian_2d_rotated(
                Xp, Yp,
                mux=row['mux'], muy=row['muy'],
                sigmax=row['sigmax'], sigmay=row['sigmay'],
                theta=row['theta_degrees'],
                amplitude=row['weight'],
                normalize=True,
                degrees=True,
            )
    radial_energy = np.sum(Zp * R, axis=0) * dtheta
    cumulative = np.cumsum(radial_energy * dr)
    total_energy = cumulative[-1] if cumulative.size else 1.0
    frac = cumulative / total_energy if total_energy > 0 else cumulative
    # Use local cubic interpolation for the 50% radius, then return HEW diameter
    radius50 = np.interp(0.5, frac, r)
    hew_m = float(2.0 * radius50)
    m_to_arcsec = 1.0 / (12.0 * np.pi / 180.0 / 3600.0)
    hew_arcsec = hew_m * m_to_arcsec
    print(f"HEW = {hew_arcsec:.4f} arcsec")

if __name__ == "__main__":
    main()
