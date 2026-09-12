"""Generalized-droop GSNR model for power-limited submarine optical links.

The implementation follows the equations used in the accompanying SubOptic 2025
paper, while adding input validation, explicit units, integer WDM channel counts,
and optional dB/km-to-linear crosstalk conversion.
"""

from __future__ import annotations

import numpy as np

from .help_functions import dB2Lin

PLANCK = 6.62607015e-34  # J s
C_NM_PER_S = 2.99792458e17  # nm/s
DB_PER_NEPER = 4.343


def _as_positive_array(name: str, value: float | np.ndarray) -> np.ndarray:
    arr = np.asarray(value, dtype=float)
    if np.any(~np.isfinite(arr)) or np.any(arr <= 0):
        raise ValueError(f"{name} must contain finite, strictly positive values.")
    return arr


def _channel_count(bandwidth_hz: float, spacing_hz: float) -> int:
    """Return the number of fully occupied WDM channels in the stated bandwidth."""
    if bandwidth_hz <= 0 or spacing_hz <= 0:
        raise ValueError("Bandwidth and channel spacing must be positive.")
    n_channels = int(np.floor(bandwidth_hz / spacing_hz + 1e-12))
    if n_channels < 1:
        raise ValueError("Bandwidth must contain at least one complete WDM channel.")
    return n_channels


def _crosstalk_linear_per_km(gamma_xt: float, unit: str) -> float:
    if unit == "linear_per_km":
        value = float(gamma_xt)
    elif unit == "dB_per_km":
        value = float(dB2Lin(gamma_xt))
    else:
        raise ValueError("gamma_xt_unit must be 'linear_per_km' or 'dB_per_km'.")
    if value < 0 or not np.isfinite(value):
        raise ValueError("Crosstalk coefficient must be finite and non-negative.")
    return value


def SNR_GD(
    Ps: float | np.ndarray,
    lamb: float,
    alpha: float,
    NF: float,
    N: int,
    gamma: float,
    Ltot: float,
    B: float,
    D: float,
    Rs: float,
    gamma_XT: float,
    l_FIFO: float,
    delta_f: float,
    *,
    gamma_xt_unit: str = "linear_per_km",
) -> np.ndarray:
    """Calculate end-to-end generalized-droop SNR in linear units.

    Parameters
    ----------
    Ps:
        Total WDM launch power per amplifier [W]. Scalars and arrays are accepted.
    lamb:
        Centre wavelength [nm].
    alpha:
        Fibre attenuation [dB/km].
    NF:
        Amplifier noise figure [dB].
    N:
        Number of modeled identical spans. Must be a positive integer.
    gamma:
        Kerr nonlinearity coefficient [1/(W km)].
    Ltot:
        Total link length [km].
    B:
        Occupied optical bandwidth [Hz].
    D:
        Chromatic dispersion parameter [s/(km nm)].
    Rs:
        Symbol rate [baud].
    gamma_XT:
        Inter-core crosstalk coefficient. Unit is selected by ``gamma_xt_unit``.
    l_FIFO:
        Insertion loss of one FIFO [dB]. Two FIFOs are included per span.
    delta_f:
        WDM channel spacing [Hz].
    gamma_xt_unit:
        ``"linear_per_km"`` or ``"dB_per_km"``.
    """
    power_w = _as_positive_array("Ps", Ps)
    if not isinstance(N, (int, np.integer)) or N < 1:
        raise ValueError("N must be a positive integer number of spans.")
    for name, value in {
        "lamb": lamb,
        "alpha": alpha,
        "gamma": gamma,
        "Ltot": Ltot,
        "B": B,
        "Rs": Rs,
        "delta_f": delta_f,
    }.items():
        if not np.isfinite(value) or value <= 0:
            raise ValueError(f"{name} must be finite and strictly positive.")
    if not np.isfinite(NF) or not np.isfinite(D) or not np.isfinite(l_FIFO):
        raise ValueError("NF, D and l_FIFO must be finite.")
    if l_FIFO < 0:
        raise ValueError("FIFO loss cannot be negative.")

    centre_frequency_hz = C_NM_PER_S / lamb
    n_channels = _channel_count(B, delta_f)
    channel_power_w = power_w / n_channels

    alpha_linear_per_km = alpha / DB_PER_NEPER
    noise_figure_linear = float(dB2Lin(NF))
    beta_2_s2_per_km = -D * lamb**2 / (2.0 * np.pi * C_NM_PER_S)
    if beta_2_s2_per_km == 0.0:
        raise ValueError("The closed-form GN expression is singular at beta_2 = 0.")

    span_length_km = Ltot / N

    # ASE and additive droop.
    span_loss_db = alpha * span_length_km + 2.0 * l_FIFO
    gain_linear = float(dB2Lin(span_loss_db))
    ase_power_w = (
        PLANCK
        * centre_frequency_hz
        * noise_figure_linear
        * Rs
        * (gain_linear - 1.0)
    )
    chi_a = 1.0 / (1.0 + ase_power_w / channel_power_w)

    # Closed-form GN-model nonlinear coefficient used in the paper.
    effective_length_km = (
        1.0 - np.exp(-alpha_linear_per_km * span_length_km)
    ) / alpha_linear_per_km
    asymptotic_effective_length_km = 1.0 / alpha_linear_per_km
    abs_beta_2 = abs(beta_2_s2_per_km)

    asinh_argument = (
        np.pi**2
        / 2.0
        * abs_beta_2
        * asymptotic_effective_length_km
        * Rs**2
        * n_channels ** (2.0 * Rs / delta_f)
    )
    alpha_nli = (
        8.0
        / 27.0
        * gamma**2
        * effective_length_km**2
        * np.arcsinh(asinh_argument)
        / (
            np.pi
            * abs_beta_2
            * asymptotic_effective_length_km
            * Rs**2
        )
    )

    inverse_snr_nli = alpha_nli * channel_power_w**2
    gamma_xt_linear = _crosstalk_linear_per_km(gamma_XT, gamma_xt_unit)
    inverse_snr_xt = gamma_xt_linear * span_length_km
    inverse_snr_redistribution = inverse_snr_nli + inverse_snr_xt
    chi_r = 1.0 / (1.0 + inverse_snr_redistribution)

    # Numerically stable equivalent of 1 / ((chi_a^-1 chi_r^-1)^N - 1).
    log_inverse_droop = -np.log(chi_a) - np.log(chi_r)
    denominator = np.expm1(N * log_inverse_droop)
    return 1.0 / denominator


# Backward-compatible lowercase alias for new code.
snr_gd = SNR_GD
