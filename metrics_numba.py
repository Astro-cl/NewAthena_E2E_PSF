try:
    from numba import njit, prange
    import numpy as np
    NUMBA_AVAILABLE = True
except Exception:
    NUMBA_AVAILABLE = False


if NUMBA_AVAILABLE:
    @njit(parallel=True, cache=True)
    def sum_gaussians(Xp, Yp, mux, muy, sigx, sigy, theta_deg, weight, normalize_flag):
        # Xp, Yp are 2D arrays (theta x r) of coordinates; others are 1D arrays
        out = np.zeros_like(Xp, dtype=np.float64)
        n_g = mux.shape[0]
        two_pi = 2.0 * np.pi
        for i in prange(n_g):
            mu_x = mux[i]
            mu_y = muy[i]
            sx = sigx[i]
            sy = sigy[i]
            th = theta_deg[i] * (np.pi / 180.0)
            c = np.cos(th)
            s = np.sin(th)
            # compute rotated coordinates
            Xr = (Xp - mu_x) * c + (Yp - mu_y) * s
            Yr = -(Xp - mu_x) * s + (Yp - mu_y) * c
            # exponent
            a = (Xr * Xr) / (2.0 * sx * sx) + (Yr * Yr) / (2.0 * sy * sy)
            if normalize_flag:
                norm = 1.0 / (two_pi * sx * sy)
            else:
                norm = 1.0
            out += weight[i] * norm * np.exp(-a)
        return out

    @njit(parallel=True)
    def radial_profile_numba(cx, cy, mux, muy, sigx, sigy, theta_deg, weight, n_r, n_theta, r_max, normalize_flag):
        """Numba-accelerated polar radial profile for mixtures of Gaussians.

        Returns (r_array, cumulative_array, total_energy)
        """
        two_pi = 2.0 * np.pi
        # build theta and r arrays
        theta = np.empty(n_theta, dtype=np.float64)
        for j in range(n_theta):
            theta[j] = two_pi * j / n_theta

        r = np.empty(n_r, dtype=np.float64)
        if n_r > 1:
            for k in range(n_r):
                r[k] = r_max * k / (n_r - 1)
        else:
            r[0] = r_max

        # allocate grids
        Xp = np.empty((n_theta, n_r), dtype=np.float64)
        Yp = np.empty((n_theta, n_r), dtype=np.float64)
        R = np.empty((n_theta, n_r), dtype=np.float64)

        for j in prange(n_theta):
            ct = np.cos(theta[j])
            st = np.sin(theta[j])
            for k in range(n_r):
                rr = r[k]
                Xp[j, k] = cx + rr * ct
                Yp[j, k] = cy + rr * st
                R[j, k] = rr

        # Evaluate Gaussian mixture on the polar grid
        Zp = sum_gaussians(Xp, Yp, mux, muy, sigx, sigy, theta_deg, weight, normalize_flag)

        dtheta = two_pi / n_theta
        dr = r[1] - r[0] if n_r > 1 else r_max

        radial_energy = np.zeros(n_r, dtype=np.float64)
        for k in range(n_r):
            s = 0.0
            for j in range(n_theta):
                s += Zp[j, k] * R[j, k]
            radial_energy[k] = s * dtheta

        cumulative = np.empty(n_r, dtype=np.float64)
        s = 0.0
        for k in range(n_r):
            s += radial_energy[k] * dr
            cumulative[k] = s

        total_energy = cumulative[n_r - 1] if n_r > 0 else 0.0
        return r, cumulative, total_energy
