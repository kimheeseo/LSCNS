"""Direct numerical helpers for the Poggiolini GN-model reference formula (GNRF).

Conventions
-----------
- Frequency: THz (= 1/ps)
- beta2: ps^2/km
- distance: km
- gamma: 1/(W km)
- alpha_db_per_km: conventional POWER attenuation in dB/km
- internally alpha_field = alpha_db_per_km * ln(10) / 20, so field ~ exp(-alpha z)
- channel power / PSD are dual-polarization totals; GNRF prefactor = 16/27
"""
from __future__ import annotations
import math
import numpy as np
from scipy.stats import qmc

C_NM_PER_PS = 299792.458
DP_GN_COEFF = 16.0 / 27.0


def alpha_field_from_db(alpha_db_per_km: float) -> float:
    return float(alpha_db_per_km) * math.log(10.0) / 20.0


def beta2_from_D(D_ps_nm_km: float, wavelength_nm: float = 1550.0) -> float:
    """Convert D [ps/(nm km)] to beta2 [ps^2/km]."""
    return -float(D_ps_nm_km) * float(wavelength_nm) ** 2 / (2.0 * math.pi * C_NM_PER_PS)


def phase_mismatch_beta2(f1_THz, f2_THz, f_THz, beta2_ps2_km):
    f1 = np.asarray(f1_THz, dtype=float)
    f2 = np.asarray(f2_THz, dtype=float)
    f = np.asarray(f_THz, dtype=float)
    return 4.0 * np.pi**2 * beta2_ps2_km * (f1 - f) * (f2 - f)


def effective_length(alpha_field_per_km: float, span_km: float) -> float:
    a = float(alpha_field_per_km)
    L = float(span_km)
    if abs(a) < 1e-15:
        return L
    return (1.0 - math.exp(-2.0 * a * L)) / (2.0 * a)


def span_efficiency_complex(delta_beta_per_km, alpha_field_per_km: float, span_km: float):
    """Literal Eq.-form magnitude squared of the single-span z integral."""
    q = np.asarray(delta_beta_per_km, dtype=float)
    a = float(alpha_field_per_km)
    L = float(span_km)
    h = (1.0 - np.exp((-2.0 * a + 1j * q) * L)) / (2.0 * a - 1j * q)
    return np.abs(h) ** 2


def span_efficiency_stable(delta_beta_per_km, alpha_field_per_km: float, span_km: float):
    """Algebraically equivalent real-valued stable form."""
    q = np.asarray(delta_beta_per_km, dtype=float)
    a = float(alpha_field_per_km)
    L = float(span_km)
    e = math.exp(-2.0 * a * L)
    return (1.0 + e**2 - 2.0 * e * np.cos(q * L)) / ((2.0 * a) ** 2 + q**2)


def phased_array_factor(delta_beta_per_km, span_km: float, n_spans: int):
    q = np.asarray(delta_beta_per_km, dtype=float)
    L = float(span_km)
    N = int(n_spans)
    x = q * L / 2.0
    den = np.sin(x)
    out = np.empty_like(x, dtype=float) if np.ndim(x) else None
    if np.ndim(x) == 0:
        if abs(float(den)) < 1e-12:
            return float(N * N)
        return float((math.sin(N * float(x)) / float(den)) ** 2)
    mask = np.abs(den) < 1e-12
    out[mask] = float(N * N)
    out[~mask] = (np.sin(N * x[~mask]) / den[~mask]) ** 2
    return out


def gn_kernel(delta_beta_per_km, alpha_field_per_km: float, span_km: float, n_spans: int):
    return span_efficiency_stable(delta_beta_per_km, alpha_field_per_km, span_km) * phased_array_factor(
        delta_beta_per_km, span_km, n_spans
    )


def rectangular_psd(power_W: float, bandwidth_THz: float) -> float:
    return float(power_W) / float(bandwidth_THz)


def gn_integrand_rectangular_point(
    f1_THz: float,
    f2_THz: float,
    f_THz: float,
    power_W: float,
    bandwidth_THz: float,
    beta2_ps2_km: float,
    gamma_W_inv_km: float,
    alpha_field_per_km: float,
    span_km: float,
    n_spans: int,
) -> float:
    """Pointwise integrand for a single rectangular channel.

    Returns 0 if f1, f2, or f3=f1+f2-f lies outside the rectangular channel.
    The remaining df1 df2 integration is NOT performed here.
    """
    B = float(bandwidth_THz)
    f3 = f1_THz + f2_THz - f_THz
    if max(abs(f1_THz), abs(f2_THz), abs(f3)) > B / 2.0:
        return 0.0
    G = rectangular_psd(power_W, B)
    q = phase_mismatch_beta2(f1_THz, f2_THz, f_THz, beta2_ps2_km)
    K = gn_kernel(q, alpha_field_per_km, span_km, n_spans)
    return float(DP_GN_COEFF * gamma_W_inv_km**2 * G**3 * K)


def gn_psd_center_single_rect_qmc(
    power_W: float = 1e-3,
    baud_GBd: float = 32.0,
    alpha_db_per_km: float = 0.2,
    D_ps_nm_km: float = 17.0,
    wavelength_nm: float = 1550.0,
    gamma_W_inv_km: float = 1.3,
    span_km: float = 80.0,
    n_spans: int = 10,
    sobol_power: int = 18,
    seed: int = 1,
) -> float:
    """2-D Sobol evaluation of G_NLI(f=0) for one rectangular channel.

    This is a compact numerical-convergence test of the GNRF double integral.
    Frequencies are integrated in THz, so the returned PSD unit is W/THz.
    """
    B = baud_GBd / 1000.0
    alpha = alpha_field_from_db(alpha_db_per_km)
    beta2 = beta2_from_D(D_ps_nm_km, wavelength_nm)
    sob = qmc.Sobol(d=2, scramble=True, seed=seed)
    u = sob.random_base2(int(sobol_power))
    f1 = (u[:, 0] - 0.5) * B
    f2 = (u[:, 1] - 0.5) * B
    f3 = f1 + f2
    valid = np.abs(f3) <= B / 2.0
    q = phase_mismatch_beta2(f1, f2, 0.0, beta2)
    K = gn_kernel(q, alpha, span_km, n_spans)
    G = power_W / B
    integrand = DP_GN_COEFF * gamma_W_inv_km**2 * G**3 * K * valid
    return float(np.mean(integrand) * B**2)


def relative_error_pct(calculated, reference) -> float:
    return float(abs(calculated - reference) / abs(reference) * 100.0)


def mathematical_self_test() -> dict:
    """Return core equation/limit/scaling checks for a reproducible example."""
    alpha = alpha_field_from_db(0.2)
    beta2 = beta2_from_D(17.0, 1550.0)
    L = 80.0
    N = 10
    Leff = effective_length(alpha, L)

    # Literal-complex vs algebraically stable single-span factor.
    q_test = np.array([-0.4, -0.1, 0.0, 0.07, 0.25])
    hc = span_efficiency_complex(q_test, alpha, L)
    hs = span_efficiency_stable(q_test, alpha, L)
    eq_residual = np.max(np.abs(hc - hs) / np.maximum(np.abs(hc), 1e-300)) * 100.0

    # delta-beta -> 0 limits.
    span0 = float(span_efficiency_stable(0.0, alpha, L))
    phase0 = phased_array_factor(0.0, L, N)
    span_limit_err = relative_error_pct(span0, Leff**2)
    phase_limit_err = relative_error_pct(phase0, N**2)

    # Cubic power scaling at a nonzero phase-mismatch point.
    kwargs = dict(
        f1_THz=0.016, f2_THz=-0.016, f_THz=0.0, bandwidth_THz=0.032,
        beta2_ps2_km=beta2, gamma_W_inv_km=1.3,
        alpha_field_per_km=alpha, span_km=L, n_spans=N,
    )
    i1 = gn_integrand_rectangular_point(power_W=1e-3, **kwargs)
    i2 = gn_integrand_rectangular_point(power_W=2e-3, **kwargs)
    scaling_ratio = i2 / i1
    scaling_err = relative_error_pct(scaling_ratio, 8.0)

    return {
        'alpha_field_1_per_km': alpha,
        'beta2_ps2_per_km': beta2,
        'Leff_km': Leff,
        'equation_residual_pct': eq_residual,
        'span_limit_error_pct': span_limit_err,
        'phase_limit_error_pct': phase_limit_err,
        'power_cubic_ratio': scaling_ratio,
        'power_cubic_error_pct': scaling_err,
    }
