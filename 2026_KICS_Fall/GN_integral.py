"""Numerical GN-model reference implementation.

This module evaluates the dual-polarization GN-model double integral instead
of using the closed-form ``asinh`` approximation in ``main.py``.

Reference
---------
P. Poggiolini et al., "The GN Model of Non-Linear Propagation in
Uncompensated Coherent Optical Systems", arXiv:1209.0394v13,
Eq. (96) and Appendix G, Eq. (G.4).

Key conventions
---------------
* Channel launch power and transmit PSD are **total dual-polarization** values.
* The full integral therefore uses the paper's 16/27 prefactor.  It must not
  be mixed with the 8/27 closed-form convention used by ``main.py``.
* Default accumulation is the finite-span coherent phased-array term from
  Eq. (96).  ``coherent_accumulation=False`` selects the faster incoherent
  N-span approximation explicitly.
* The GN model assumes Gaussian-distributed symbols.  Selecting BPSK/QPSK/
  QAM below changes the recorded rate and bandwidth inputs, but does not by
  itself make this an EGN or finite-constellation model.  Use an independently
  justified ``nli_correction_factor`` for that purpose; it is never fitted
  automatically.

The core calculation is intentionally slower than a closed form: it performs
Gauss-Legendre integration over f1, f2, and the channel-under-test bandwidth.
Use ``convergence_study`` to choose the quadrature order needed for a target
accuracy before executing a large launch-power sweep.

Example
-------
    from GN_integral import (
        FiberParameters, SystemParameters, GNIntegralOptions,
        simulate_snr_curve,
    )

    fiber = FiberParameters(
        name="G.654.E", attenuation_db_per_km=0.166,
        effective_area_um2=125.0, dispersion_ps_nm_km=21.0,
        n2_m2_w=2.2e-20,
    )
    system = SystemParameters(
        channels=90, symbol_rate_gbd=95.0, spacing_ghz=95.0,
        span_length_km=80.0, noise_figure_db=5.0,
        transceiver_snr_db=18.0,
    )
    result = simulate_snr_curve(
        launch_dbm=[-10, -5, 0, 4, 8, 10], spans=30,
        fiber=fiber, system=system,
        options=GNIntegralOptions(quadrature_order=24,
                                  output_quadrature_order=5),
        modulation="64QAM",
    )
    print(result["snr_db"])

CLI example
-----------
    python GN_integral.py --channels 90 --spans 30 --modulation 64QAM \
        --quadrature-order 16 --output-quadrature-order 5 \
        --save-csv gn_integral_curve.csv
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import time
from dataclasses import asdict, dataclass, replace
from pathlib import Path
from typing import Any, Dict, Iterable, Mapping, Optional, Sequence, Tuple

import numpy as np


C0 = 299_792_458.0
H_PLANCK = 6.626_070_15e-34
DB_PER_NEPER_POWER = 10.0 / np.log(10.0)
PAPER_REFERENCE = "Poggiolini et al., arXiv:1209.0394v13, Eq. (96)/(G.4)"


# ---------------------------------------------------------------------------
# Input models
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class FiberParameters:
    """Fiber inputs, using the same base units as ``main.py``.

    Direct alpha/beta2/gamma values override the values derived from usual
    catalog inputs.  ``beta3_s3_per_km`` is optional and defaults to zero;
    it implements the beta2 + beta3 phase mismatch in Appendix G.
    """

    name: str = "G.654.E"
    wavelength_nm: float = 1550.0
    attenuation_db_per_km: float = 0.166
    effective_area_um2: float = 125.0
    dispersion_ps_nm_km: float = 21.0
    n2_m2_w: float = 2.2e-20
    alpha_power_per_km: Optional[float] = None
    beta2_s2_per_km: Optional[float] = None
    beta3_s3_per_km: float = 0.0
    gamma_per_w_km: Optional[float] = None


@dataclass(frozen=True)
class SystemParameters:
    """WDM, span, amplifier, and transceiver inputs.

    ``ase_bandwidth_hz=None`` keeps the convention in ``main.py``: ASE is
    integrated over the symbol rate.  The numerical NLI, in contrast, is
    integrated over the actual channel support including roll-off.
    """

    channels: int = 90
    symbol_rate_gbd: float = 95.0
    spacing_ghz: float = 95.0
    span_length_km: float = 80.0
    noise_figure_db: float = 5.0
    transceiver_snr_db: float = 18.0
    polarizations: int = 2
    ase_bandwidth_hz: Optional[float] = None
    stated_gain_bandwidth_thz: Optional[float] = 4.8
    qam_order_record_only: int = 64
    shannon_gap_db_record_only: float = 3.0


@dataclass(frozen=True)
class WDMChannel:
    """One WDM channel with a unit-area rectangular or raised-cosine PSD.

    ``power_w`` is total DP channel power.  A raised-cosine spectrum is useful
    when changing roll-off; a zero roll-off is the exact rectangular Nyquist
    PSD normally assumed by the GN closed form.
    """

    center_frequency_hz: float
    symbol_rate_hz: float
    power_w: float
    roll_off: float = 0.0
    modulation: str = "64QAM"
    name: str = ""

    def __post_init__(self) -> None:
        if not np.isfinite(self.center_frequency_hz):
            raise ValueError("center_frequency_hz must be finite.")
        if not np.isfinite(self.symbol_rate_hz) or self.symbol_rate_hz <= 0:
            raise ValueError("symbol_rate_hz must be positive and finite.")
        if not np.isfinite(self.power_w) or self.power_w < 0:
            raise ValueError("power_w must be non-negative and finite.")
        if not np.isfinite(self.roll_off) or not 0.0 <= self.roll_off <= 1.0:
            raise ValueError("roll_off must be between 0 and 1.")

    @property
    def bandwidth_hz(self) -> float:
        """Occupied support width, R_s(1 + roll_off), in Hz."""

        return self.symbol_rate_hz * (1.0 + self.roll_off)

    @property
    def lower_frequency_hz(self) -> float:
        return self.center_frequency_hz - 0.5 * self.bandwidth_hz

    @property
    def upper_frequency_hz(self) -> float:
        return self.center_frequency_hz + 0.5 * self.bandwidth_hz

    def normalized_psd(self, frequency_hz: np.ndarray | float) -> np.ndarray:
        """Return a unit-area PSD in 1/Hz at the supplied relative frequency."""

        frequency = np.asarray(frequency_hz, dtype=float)
        offset = np.abs(frequency - self.center_frequency_hz)
        rate = self.symbol_rate_hz
        roll = self.roll_off
        if roll == 0.0:
            return np.where(offset <= 0.5 * rate, 1.0 / rate, 0.0)

        flat_edge = 0.5 * rate * (1.0 - roll)
        outer_edge = 0.5 * rate * (1.0 + roll)
        result = np.zeros_like(offset, dtype=float)
        flat = offset <= flat_edge
        transition = (offset > flat_edge) & (offset <= outer_edge)
        result[flat] = 1.0 / rate
        if np.any(transition):
            argument = np.pi * (offset[transition] - flat_edge) / (roll * rate)
            result[transition] = 0.5 * (1.0 + np.cos(argument)) / rate
        return result

    def psd(self, frequency_hz: np.ndarray | float) -> np.ndarray:
        """Return the total-DP transmit PSD in W/Hz."""

        return self.power_w * self.normalized_psd(frequency_hz)


@dataclass(frozen=True)
class GNIntegralOptions:
    """Accuracy and physical-convention switches for the numerical integral.

    ``quadrature_order`` applies in each f1/f2 dimension.  Increasing it is
    the primary accuracy/runtime trade-off.  ``output_quadrature_order`` also
    integrates the output channel bandwidth, rather than using only its center.
    The default coherent option is the paper's finite-span phased-array term.
    ``phase_matching_refinement`` adds geometric subintervals only around the
    f1=f and f2=f phase-matching ridges, avoiding the poor convergence caused
    by stepping across those narrow SCI/XCI features.
    """

    quadrature_order: int = 24
    output_quadrature_order: int = 3
    coherent_accumulation: bool = True
    nli_prefactor: float = 16.0 / 27.0
    nli_correction_factor: float = 1.0
    power_definition: str = "total_dp"
    phase_matching_refinement: int = 5

    def __post_init__(self) -> None:
        if self.quadrature_order < 2:
            raise ValueError("quadrature_order must be at least 2.")
        if self.output_quadrature_order < 1:
            raise ValueError("output_quadrature_order must be at least 1.")
        if not np.isfinite(self.nli_prefactor) or self.nli_prefactor <= 0:
            raise ValueError("nli_prefactor must be positive and finite.")
        if not np.isfinite(self.nli_correction_factor) or self.nli_correction_factor <= 0:
            raise ValueError("nli_correction_factor must be positive and finite.")
        if (
            isinstance(self.phase_matching_refinement, bool)
            or int(self.phase_matching_refinement) != self.phase_matching_refinement
            or self.phase_matching_refinement < 0
            or self.phase_matching_refinement > 12
        ):
            raise ValueError("phase_matching_refinement must be an integer from 0 to 12.")
        if self.power_definition != "total_dp":
            raise ValueError(
                "This implementation uses total dual-polarization channel power; "
                "set power_definition='total_dp'."
            )


@dataclass(frozen=True)
class GNIntegralResult:
    """Numerical NLI result for a channel under test (CUT)."""

    coi_index: int
    nli_power_w: float
    eta_equivalent_w_inv2: float
    output_frequencies_hz: np.ndarray
    output_psd_w_per_hz: np.ndarray
    metadata: Mapping[str, Any]

    def as_dict(self) -> Dict[str, Any]:
        return {
            "coi_index": self.coi_index,
            "nli_power_w": self.nli_power_w,
            "eta_equivalent_w_inv2": self.eta_equivalent_w_inv2,
            "output_frequencies_hz": self.output_frequencies_hz.copy(),
            "output_psd_w_per_hz": self.output_psd_w_per_hz.copy(),
            "metadata": dict(self.metadata),
        }


@dataclass
class _IntegrationStats:
    triplet_domains: int = 0
    quadrature_points: int = 0


# ---------------------------------------------------------------------------
# Unit conversion, modulation helpers, and WDM-plan construction
# ---------------------------------------------------------------------------


def dbm_to_w(power_dbm: float | Sequence[float] | np.ndarray) -> np.ndarray:
    """Convert dBm to W."""

    return 1e-3 * np.power(10.0, np.asarray(power_dbm, dtype=float) / 10.0)


def w_to_dbm(power_w: float | Sequence[float] | np.ndarray) -> np.ndarray:
    """Convert a strictly positive optical power in W to dBm."""

    power = np.asarray(power_w, dtype=float)
    if np.any(power <= 0):
        raise ValueError("Power must be strictly positive to convert to dBm.")
    return 10.0 * np.log10(power / 1e-3)


def modulation_bits_per_symbol(modulation: str) -> float:
    """Return bits/symbol/polarization from common modulation labels.

    Labels may include a ``PM-`` or ``DP-`` prefix; DP rate is handled later
    using the explicit ``polarizations`` input rather than double-counted here.
    """

    normalized = re.sub(r"[^A-Z0-9]", "", modulation.upper())
    normalized = normalized.removeprefix("PM").removeprefix("DP")
    known = {
        "BPSK": 1.0,
        "QPSK": 2.0,
        "8QAM": 3.0,
        "16QAM": 4.0,
        "32QAM": 5.0,
        "64QAM": 6.0,
        "128QAM": 7.0,
        "256QAM": 8.0,
        "512QAM": 9.0,
        "1024QAM": 10.0,
    }
    if normalized in known:
        return known[normalized]
    match = re.fullmatch(r"(\d+)QAM", normalized)
    if match:
        order = int(match.group(1))
        bits = np.log2(order)
        if order > 1 and np.isclose(bits, round(bits)):
            return float(round(bits))
    raise ValueError(f"Unsupported modulation label: {modulation!r}")


def _broadcast(value: Any, length: int, name: str) -> list[Any]:
    if isinstance(value, str) or np.isscalar(value):
        return [value] * length
    result = list(value)
    if len(result) != length:
        raise ValueError(f"{name} must be scalar or have exactly {length} entries.")
    return result


def make_uniform_wdm(
    system: SystemParameters,
    *,
    launch_power_dbm: float | Sequence[float] | np.ndarray = 0.0,
    roll_off: float | Sequence[float] = 0.0,
    modulation: str | Sequence[str] = "64QAM",
    channel_power_offsets_db: Optional[Sequence[float]] = None,
) -> Tuple[WDMChannel, ...]:
    """Construct a symmetric WDM plan from ``SystemParameters``.

    For an even number of channels there is no exact zero-frequency channel;
    channel ``channels // 2`` is the default channel under test.  All frequency
    coordinates are offsets from the common expansion/reference frequency, so
    an absolute optical carrier is not required for beta2-only calculations.
    """

    count = int(system.channels)
    if count < 1 or count != system.channels:
        raise ValueError("system.channels must be a positive integer.")
    if system.symbol_rate_gbd <= 0 or system.spacing_ghz <= 0:
        raise ValueError("symbol_rate_gbd and spacing_ghz must be positive.")

    power_values = _broadcast(launch_power_dbm, count, "launch_power_dbm")
    if channel_power_offsets_db is not None:
        offsets = _broadcast(channel_power_offsets_db, count, "channel_power_offsets_db")
        power_values = [float(power) + float(offset) for power, offset in zip(power_values, offsets)]
    rolls = _broadcast(roll_off, count, "roll_off")
    modulations = _broadcast(modulation, count, "modulation")
    rate = float(system.symbol_rate_gbd) * 1e9
    spacing = float(system.spacing_ghz) * 1e9
    centers = (np.arange(count, dtype=float) - 0.5 * (count - 1)) * spacing

    return tuple(
        WDMChannel(
            center_frequency_hz=float(center),
            symbol_rate_hz=rate,
            power_w=float(dbm_to_w(float(power))),
            roll_off=float(roll),
            modulation=str(mod),
            name=f"ch{index}",
        )
        for index, (center, power, roll, mod) in enumerate(
            zip(centers, power_values, rolls, modulations)
        )
    )


def default_coi_index(channels: Sequence[WDMChannel]) -> int:
    """Return the central/right-central index used by the uniform-plan helper."""

    if not channels:
        raise ValueError("At least one WDM channel is required.")
    return len(channels) // 2


def _object_value(obj: Any, name: str, default: Any = None) -> Any:
    return getattr(obj, name, default)


def _object_to_dict(obj: Any) -> Dict[str, Any]:
    try:
        return asdict(obj)
    except TypeError:
        return {
            name: _object_value(obj, name)
            for name in dir(obj)
            if not name.startswith("_") and not callable(_object_value(obj, name))
        }


def derive_link_parameters(
    fiber: FiberParameters | Any = FiberParameters(),
    system: SystemParameters | Any = SystemParameters(),
    *,
    beta3_s3_per_km: Optional[float] = None,
) -> Dict[str, float]:
    """Derive alpha, beta2, beta3, gamma, and ASE quantities with explicit units.

    A ``FiberParameters``/``SystemParameters`` object imported from ``main.py``
    is accepted as well.  This makes the numerical file a drop-in companion to
    the existing closed-form code without importing ``main.py`` at runtime.
    """

    wavelength_nm = float(_object_value(fiber, "wavelength_nm", 1550.0))
    attenuation_db_per_km = float(_object_value(fiber, "attenuation_db_per_km", 0.166))
    effective_area_um2 = float(_object_value(fiber, "effective_area_um2", 125.0))
    dispersion_ps_nm_km = float(_object_value(fiber, "dispersion_ps_nm_km", 21.0))
    n2_m2_w = float(_object_value(fiber, "n2_m2_w", 2.2e-20))
    alpha_override = _object_value(fiber, "alpha_power_per_km", None)
    beta2_override = _object_value(fiber, "beta2_s2_per_km", None)
    gamma_override = _object_value(fiber, "gamma_per_w_km", None)
    beta3_value = (
        _object_value(fiber, "beta3_s3_per_km", 0.0)
        if beta3_s3_per_km is None
        else beta3_s3_per_km
    )

    symbol_rate_gbd = float(_object_value(system, "symbol_rate_gbd", 95.0))
    spacing_ghz = float(_object_value(system, "spacing_ghz", 95.0))
    span_length_km = float(_object_value(system, "span_length_km", 80.0))
    noise_figure_db = float(_object_value(system, "noise_figure_db", 5.0))
    ase_bandwidth_override = _object_value(system, "ase_bandwidth_hz", None)

    for name, value in {
        "wavelength_nm": wavelength_nm,
        "effective_area_um2": effective_area_um2,
        "n2_m2_w": n2_m2_w,
        "symbol_rate_gbd": symbol_rate_gbd,
        "spacing_ghz": spacing_ghz,
        "span_length_km": span_length_km,
    }.items():
        if not np.isfinite(value) or value <= 0:
            raise ValueError(f"{name} must be positive and finite.")
    if not np.isfinite(noise_figure_db):
        raise ValueError("noise_figure_db must be finite.")
    if not np.isfinite(float(beta3_value)):
        raise ValueError("beta3_s3_per_km must be finite.")

    wavelength_m = wavelength_nm * 1e-9
    rate_hz = symbol_rate_gbd * 1e9
    spacing_hz = spacing_ghz * 1e9
    if alpha_override is None:
        if attenuation_db_per_km <= 0:
            raise ValueError("attenuation_db_per_km must be positive.")
        alpha_power_per_km = attenuation_db_per_km / DB_PER_NEPER_POWER
    else:
        alpha_power_per_km = float(alpha_override)
        if alpha_power_per_km <= 0:
            raise ValueError("alpha_power_per_km must be positive.")
        attenuation_db_per_km = DB_PER_NEPER_POWER * alpha_power_per_km

    if beta2_override is None:
        # 1 ps/(nm km) = 1e-6 s/m^2; convert beta2 from D at wavelength.
        dispersion_si = dispersion_ps_nm_km * 1e-6
        beta2_s2_per_km = -(wavelength_m**2 / (2.0 * np.pi * C0)) * dispersion_si * 1000.0
    else:
        beta2_s2_per_km = float(beta2_override)
    if beta2_s2_per_km == 0.0:
        raise ValueError("beta2_s2_per_km cannot be zero in this GN implementation.")

    if gamma_override is None:
        gamma_per_w_km = (
            2.0 * np.pi * n2_m2_w / (wavelength_m * effective_area_um2 * 1e-12) * 1000.0
        )
    else:
        gamma_per_w_km = float(gamma_override)
    if gamma_per_w_km <= 0.0 or not np.isfinite(gamma_per_w_km):
        raise ValueError("gamma_per_w_km must be positive and finite.")

    span_gain_linear = np.exp(alpha_power_per_km * span_length_km)
    ase_bandwidth_hz = rate_hz if ase_bandwidth_override is None else float(ase_bandwidth_override)
    if ase_bandwidth_hz <= 0.0 or not np.isfinite(ase_bandwidth_hz):
        raise ValueError("ase_bandwidth_hz must be positive and finite.")
    noise_figure_linear = 10.0 ** (noise_figure_db / 10.0)
    ase_per_span_w = (
        H_PLANCK
        * (C0 / wavelength_m)
        * noise_figure_linear
        * ase_bandwidth_hz
        * (span_gain_linear - 1.0)
    )

    return {
        "wavelength_m": wavelength_m,
        "frequency_hz": C0 / wavelength_m,
        "rate_hz": rate_hz,
        "spacing_hz": spacing_hz,
        "span_length_km": span_length_km,
        # alpha_power = 2*alpha in Eq. (96)/(G.4), where alpha is field loss.
        "alpha_power_per_km": alpha_power_per_km,
        "alpha_field_per_km": alpha_power_per_km / 2.0,
        "beta2_s2_per_km": beta2_s2_per_km,
        "beta3_s3_per_km": float(beta3_value),
        "gamma_per_w_km": gamma_per_w_km,
        "effective_length_km": -np.expm1(-alpha_power_per_km * span_length_km) / alpha_power_per_km,
        "span_gain_linear": span_gain_linear,
        "span_gain_db": 10.0 * np.log10(span_gain_linear),
        "ase_bandwidth_hz": ase_bandwidth_hz,
        "ase_per_span_w": ase_per_span_w,
        "attenuation_db_per_km_used": attenuation_db_per_km,
    }


# ---------------------------------------------------------------------------
# Equation (96)/(G.4) kernel and deterministic numerical integration
# ---------------------------------------------------------------------------


def phase_mismatch_rad_per_km(
    f1_hz: np.ndarray | float,
    f2_hz: np.ndarray | float,
    f_hz: float,
    parameters: Mapping[str, float],
) -> np.ndarray:
    """Return the Appendix-G phase mismatch in rad/km.

    This is the exact beta2 + beta3 form in Eq. (G.2):

    ``4*pi^2*(f1-f)*(f2-f) * [beta2 + pi*beta3*(f1+f2)]``.
    """

    f1 = np.asarray(f1_hz, dtype=float)
    f2 = np.asarray(f2_hz, dtype=float)
    beta2 = float(parameters["beta2_s2_per_km"])
    beta3 = float(parameters["beta3_s3_per_km"])
    return 4.0 * np.pi**2 * (f1 - f_hz) * (f2 - f_hz) * (
        beta2 + np.pi * beta3 * (f1 + f2)
    )


def _phased_array_squared(phase_rad: np.ndarray, spans: int) -> np.ndarray:
    """Stable evaluation of sin(N*phase/2)^2 / sin(phase/2)^2."""

    half_phase = 0.5 * np.asarray(phase_rad, dtype=float)
    denominator = np.sin(half_phase)
    result = np.empty_like(half_phase, dtype=float)
    near_resonance = np.abs(denominator) < 1e-10
    result[near_resonance] = float(spans * spans)
    if np.any(~near_resonance):
        ratio = np.sin(spans * half_phase[~near_resonance]) / denominator[~near_resonance]
        result[~near_resonance] = ratio * ratio
    return result


def link_function_squared(
    delta_beta_rad_per_km: np.ndarray | float,
    parameters: Mapping[str, float],
    spans: int,
    *,
    coherent_accumulation: bool,
) -> np.ndarray:
    """Return the finite-span kernel in Eq. (96), with units km^2.

    The one-span factor is
    ``|(1-exp(-2*alpha*L)*exp(j*DeltaBeta*L))/(2*alpha-j*DeltaBeta)|^2``.
    The final multiplier is either the exact phased-array factor or N spans.
    """

    if spans < 1 or int(spans) != spans:
        raise ValueError("spans must be a positive integer.")
    delta = np.asarray(delta_beta_rad_per_km, dtype=float)
    length = float(parameters["span_length_km"])
    alpha_power = float(parameters["alpha_power_per_km"])
    denominator = alpha_power - 1j * delta
    one_span_field = -np.expm1(-denominator * length) / denominator
    one_span_squared = np.abs(one_span_field) ** 2
    if coherent_accumulation:
        accumulation = _phased_array_squared(delta * length, int(spans))
    else:
        accumulation = float(spans)
    return one_span_squared * accumulation


def _split_outer_intervals(
    f1_low: float,
    f1_high: float,
    f2_low: float,
    f2_high: float,
    f3_low: float,
    f3_high: float,
    f_out: float,
    phase_matching_refinement: int,
) -> np.ndarray:
    """Split f1 where the f2 intersection changes branch (improves accuracy)."""

    breakpoints = [
        f1_low,
        f1_high,
        # DeltaBeta is exactly zero at f1=f_out.  Explicitly splitting there
        # prevents Gauss nodes from stepping over the narrow high-kernel ridge.
        f_out,
        f3_low + f_out - f2_low,
        f3_low + f_out - f2_high,
        f3_high + f_out - f2_low,
        f3_high + f_out - f2_high,
    ]
    if f1_low < f_out < f1_high and phase_matching_refinement:
        fractions = _phase_matching_fractions(phase_matching_refinement)
        left_distance = f_out - f1_low
        right_distance = f1_high - f_out
        breakpoints.extend(f_out - left_distance * fractions)
        breakpoints.extend(f_out + right_distance * fractions)
    clipped = np.clip(np.asarray(breakpoints, dtype=float), f1_low, f1_high)
    return np.unique(clipped)


def _phase_matching_fractions(refinement: int) -> np.ndarray:
    """Geometric intervals that resolve the DeltaBeta=0 SCI/XCI ridges."""

    if refinement <= 0:
        return np.asarray([0.0, 1.0])
    return np.concatenate(([0.0], 10.0 ** -np.arange(refinement, 0, -1), [1.0]))


def _integrate_triplet_domain(
    channel_1: WDMChannel,
    channel_2: WDMChannel,
    channel_3: WDMChannel,
    f_out_hz: float,
    *,
    outer_nodes: np.ndarray,
    outer_weights: np.ndarray,
    parameters: Mapping[str, float],
    spans: int,
    options: GNIntegralOptions,
    stats: _IntegrationStats,
) -> float:
    """Integrate one (channel_1, channel_2, channel_3) spectral triplet."""

    f1_low, f1_high = channel_1.lower_frequency_hz, channel_1.upper_frequency_hz
    f2_low, f2_high = channel_2.lower_frequency_hz, channel_2.upper_frequency_hz
    f3_low, f3_high = channel_3.lower_frequency_hz, channel_3.upper_frequency_hz
    split = _split_outer_intervals(
        f1_low,
        f1_high,
        f2_low,
        f2_high,
        f3_low,
        f3_high,
        f_out_hz,
        options.phase_matching_refinement,
    )
    total = 0.0
    for left, right in zip(split[:-1], split[1:]):
        if right <= left:
            continue
        f1 = 0.5 * (right + left) + 0.5 * (right - left) * outer_nodes
        w1 = 0.5 * (right - left) * outer_weights
        f2_lower = np.maximum(f2_low, f3_low - f1 + f_out_hz)
        f2_upper = np.minimum(f2_high, f3_high - f1 + f_out_hz)
        valid = f2_upper > f2_lower
        if not np.any(valid):
            continue

        f1 = f1[valid]
        w1 = w1[valid]
        f2_lower = f2_lower[valid]
        f2_upper = f2_upper[valid]

        def add_piece(row_mask: np.ndarray, lower_piece: np.ndarray, upper_piece: np.ndarray) -> None:
            """Accumulate one vectorized f2 subinterval for selected f1 nodes."""

            nonlocal total
            valid_piece = row_mask & (upper_piece > lower_piece)
            if not np.any(valid_piece):
                return
            f1_piece = f1[valid_piece]
            w1_piece = w1[valid_piece]
            lo = lower_piece[valid_piece]
            hi = upper_piece[valid_piece]
            half_width = 0.5 * (hi - lo)
            f2 = 0.5 * (hi + lo)[:, None] + half_width[:, None] * outer_nodes[None, :]
            w2 = half_width[:, None] * outer_weights[None, :]
            f1_2d = f1_piece[:, None]
            f3 = f1_2d + f2 - f_out_hz
            density = channel_1.psd(f1_piece)[:, None] * channel_2.psd(f2) * channel_3.psd(f3)
            delta_beta = phase_mismatch_rad_per_km(f1_2d, f2, f_out_hz, parameters)
            kernel = link_function_squared(
                delta_beta,
                parameters,
                spans,
                coherent_accumulation=options.coherent_accumulation,
            )
            total += float(np.sum(w1_piece[:, None] * w2 * density * kernel))
            stats.quadrature_points += int(f1_piece.size * outer_nodes.size)

        crosses_output = (f2_lower < f_out_hz) & (f2_upper > f_out_hz)
        # Regions away from f2=f_out have no zero-phase ridge and can be
        # integrated as one interval.  Only the CUT-containing f2 domain is
        # geometrically refined, keeping the cost local rather than global.
        add_piece(~crosses_output, f2_lower, f2_upper)
        if np.any(crosses_output):
            fractions = _phase_matching_fractions(options.phase_matching_refinement)
            left_distance = f_out_hz - f2_lower
            right_distance = f2_upper - f_out_hz
            for start, end in zip(fractions[:-1], fractions[1:]):
                add_piece(
                    crosses_output,
                    f_out_hz - left_distance * end,
                    f_out_hz - left_distance * start,
                )
                add_piece(
                    crosses_output,
                    f_out_hz + right_distance * start,
                    f_out_hz + right_distance * end,
                )
    return total


def _nli_psd_at_frequency(
    f_out_hz: float,
    channels: Sequence[WDMChannel],
    *,
    parameters: Mapping[str, float],
    spans: int,
    options: GNIntegralOptions,
    nodes: np.ndarray,
    weights: np.ndarray,
    stats: _IntegrationStats,
) -> float:
    """Evaluate Eq. (96) at one output frequency within the CUT band."""

    lowers = np.asarray([channel.lower_frequency_hz for channel in channels], dtype=float)
    uppers = np.asarray([channel.upper_frequency_hz for channel in channels], dtype=float)
    integral = 0.0
    for channel_1 in channels:
        if channel_1.power_w == 0.0:
            continue
        for channel_2 in channels:
            if channel_2.power_w == 0.0:
                continue
            f3_min = channel_1.lower_frequency_hz + channel_2.lower_frequency_hz - f_out_hz
            f3_max = channel_1.upper_frequency_hz + channel_2.upper_frequency_hz - f_out_hz
            candidates = np.flatnonzero((uppers > f3_min) & (lowers < f3_max))
            for index in candidates:
                channel_3 = channels[int(index)]
                if channel_3.power_w == 0.0:
                    continue
                stats.triplet_domains += 1
                integral += _integrate_triplet_domain(
                    channel_1,
                    channel_2,
                    channel_3,
                    f_out_hz,
                    outer_nodes=nodes,
                    outer_weights=weights,
                    parameters=parameters,
                    spans=spans,
                    options=options,
                    stats=stats,
                )
    scale = options.nli_prefactor * parameters["gamma_per_w_km"] ** 2
    return float(scale * options.nli_correction_factor * integral)


def integrate_nli_for_channel(
    channels: Sequence[WDMChannel],
    coi_index: Optional[int] = None,
    *,
    fiber: FiberParameters | Any = FiberParameters(),
    system: SystemParameters | Any = SystemParameters(),
    spans: int = 1,
    options: GNIntegralOptions = GNIntegralOptions(),
    beta3_s3_per_km: Optional[float] = None,
) -> GNIntegralResult:
    """Integrate Eq. (96)/(G.4) across the full bandwidth of one CUT.

    The returned ``eta_equivalent_w_inv2`` is ``P_NLI/P_CUT^3`` for the
    *provided WDM power profile*.  It can be re-used for a launch sweep only
    when every channel power is scaled by the same factor.
    """

    channels = tuple(channels)
    if not channels:
        raise ValueError("At least one WDM channel is required.")
    if spans < 1 or int(spans) != spans:
        raise ValueError("spans must be a positive integer.")
    cut = default_coi_index(channels) if coi_index is None else int(coi_index)
    if not 0 <= cut < len(channels):
        raise IndexError("coi_index is outside the WDM channel plan.")
    if channels[cut].power_w <= 0.0:
        raise ValueError("The channel under test must have positive power.")

    parameters = derive_link_parameters(fiber, system, beta3_s3_per_km=beta3_s3_per_km)
    nodes, weights = np.polynomial.legendre.leggauss(options.quadrature_order)
    output_nodes, output_weights = np.polynomial.legendre.leggauss(options.output_quadrature_order)
    channel = channels[cut]
    f_low, f_high = channel.lower_frequency_hz, channel.upper_frequency_hz
    output_frequency = 0.5 * (f_high + f_low) + 0.5 * (f_high - f_low) * output_nodes
    output_weight = 0.5 * (f_high - f_low) * output_weights
    stats = _IntegrationStats()
    started = time.perf_counter()
    output_psd = np.asarray(
        [
            _nli_psd_at_frequency(
                float(frequency),
                channels,
                parameters=parameters,
                spans=int(spans),
                options=options,
                nodes=nodes,
                weights=weights,
                stats=stats,
            )
            for frequency in output_frequency
        ],
        dtype=float,
    )
    nli_power = float(np.sum(output_weight * output_psd))
    elapsed = time.perf_counter() - started
    eta = nli_power / channels[cut].power_w**3
    metadata: Dict[str, Any] = {
        "paper_reference": PAPER_REFERENCE,
        "formula": "dual-polarization numerical double integral with finite-span kernel",
        "power_definition": "total DP channel power and PSD",
        "nli_prefactor": options.nli_prefactor,
        "nli_correction_factor": options.nli_correction_factor,
        "coherent_accumulation": options.coherent_accumulation,
        "spans": int(spans),
        "quadrature_order": options.quadrature_order,
        "output_quadrature_order": options.output_quadrature_order,
        "triplet_domains": stats.triplet_domains,
        "quadrature_points": stats.quadrature_points,
        "runtime_s": elapsed,
        "cut_bandwidth_hz": channel.bandwidth_hz,
        "cut_modulation": channel.modulation,
        "derived_parameters": parameters,
    }
    return GNIntegralResult(
        coi_index=cut,
        nli_power_w=nli_power,
        eta_equivalent_w_inv2=float(eta),
        output_frequencies_hz=output_frequency,
        output_psd_w_per_hz=output_psd,
        metadata=metadata,
    )


# Clear aliases for users who prefer a verb closer to the paper terminology.
evaluate_gn_integral = integrate_nli_for_channel
compute_nli = integrate_nli_for_channel


def convergence_study(
    channels: Sequence[WDMChannel],
    coi_index: Optional[int] = None,
    *,
    quadrature_orders: Iterable[int] = (8, 12, 16, 20),
    fiber: FiberParameters | Any = FiberParameters(),
    system: SystemParameters | Any = SystemParameters(),
    spans: int = 1,
    options: GNIntegralOptions = GNIntegralOptions(),
    beta3_s3_per_km: Optional[float] = None,
) -> list[Dict[str, float]]:
    """Run successive quadrature orders and report numerical convergence.

    The last order is the numerical reference only for this convergence test;
    it is not a substitute for validation against a split-step simulation or
    measured link.  A small final relative change indicates integration error
    is controlled for the chosen physical model.
    """

    orders = [int(order) for order in quadrature_orders]
    if not orders or any(order < 2 for order in orders):
        raise ValueError("quadrature_orders must contain integers >= 2.")
    rows: list[Dict[str, float]] = []
    for order in orders:
        result = integrate_nli_for_channel(
            channels,
            coi_index,
            fiber=fiber,
            system=system,
            spans=spans,
            options=replace(options, quadrature_order=order),
            beta3_s3_per_km=beta3_s3_per_km,
        )
        rows.append(
            {
                "quadrature_order": float(order),
                "nli_power_w": result.nli_power_w,
                "eta_equivalent_w_inv2": result.eta_equivalent_w_inv2,
                "runtime_s": float(result.metadata["runtime_s"]),
            }
        )
    finest = rows[-1]["eta_equivalent_w_inv2"]
    previous: Optional[float] = None
    for row in rows:
        eta = row["eta_equivalent_w_inv2"]
        row["relative_to_finest_pct"] = 100.0 * (eta - finest) / finest
        row["absolute_relative_to_finest_pct"] = abs(row["relative_to_finest_pct"])
        row["relative_change_from_previous_pct"] = (
            float("nan") if previous is None else 100.0 * (eta - previous) / previous
        )
        previous = eta
    return rows


# ---------------------------------------------------------------------------
# SNR / throughput helpers for uniform launch-power sweeps
# ---------------------------------------------------------------------------


def _as_float_array(values: float | Sequence[float] | np.ndarray) -> np.ndarray:
    array = np.atleast_1d(np.asarray(values, dtype=float))
    if not np.all(np.isfinite(array)):
        raise ValueError("Input array contains non-finite values.")
    return array


def _format_limited_information_rate_gbps(
    snr_linear: np.ndarray,
    *,
    symbol_rate_hz: float,
    modulation: str,
    polarizations: int,
    coding_rate: float,
    shannon_gap_db: float,
) -> np.ndarray:
    """Return a transparent, format-capped Shannon estimate, not a FEC curve."""

    if not 0.0 < coding_rate <= 1.0:
        raise ValueError("coding_rate must be in (0, 1].")
    if polarizations not in (1, 2):
        raise ValueError("polarizations must be 1 or 2.")
    constellation_cap = modulation_bits_per_symbol(modulation)
    gap_linear = 10.0 ** (shannon_gap_db / 10.0)
    bits_per_symbol_per_pol = np.minimum(
        constellation_cap, np.log2(1.0 + snr_linear / gap_linear)
    )
    return polarizations * symbol_rate_hz * coding_rate * bits_per_symbol_per_pol / 1e9


def simulate_snr_curve(
    launch_dbm: float | Sequence[float] | np.ndarray,
    spans: int,
    *,
    fiber: FiberParameters | Any = FiberParameters(),
    system: SystemParameters | Any = SystemParameters(),
    options: GNIntegralOptions = GNIntegralOptions(),
    modulation: str | Sequence[str] = "64QAM",
    roll_off: float | Sequence[float] = 0.0,
    coi_index: Optional[int] = None,
    channel_power_offsets_db: Optional[Sequence[float]] = None,
    beta3_s3_per_km: Optional[float] = None,
    include_transceiver_noise: bool = True,
    coding_rate: float = 1.0,
    shannon_gap_db: Optional[float] = None,
) -> Dict[str, Any]:
    """Calculate a launch-power SNR curve with one numerical NLI integral.

    The integral is evaluated once at 0 dBm CUT power and converted to an
    equivalent eta.  This is exact for a sweep that scales **every WDM channel
    power by the same factor**, including optional fixed per-channel offsets.
    It avoids needlessly repeating a costly integral for each launch point.
    """

    if spans < 1 or int(spans) != spans:
        raise ValueError("spans must be a positive integer.")
    launch = _as_float_array(launch_dbm)
    plan = make_uniform_wdm(
        system,
        launch_power_dbm=0.0,
        roll_off=roll_off,
        modulation=modulation,
        channel_power_offsets_db=channel_power_offsets_db,
    )
    cut = default_coi_index(plan) if coi_index is None else int(coi_index)
    nli_result = integrate_nli_for_channel(
        plan,
        cut,
        fiber=fiber,
        system=system,
        spans=int(spans),
        options=options,
        beta3_s3_per_km=beta3_s3_per_km,
    )
    parameters = derive_link_parameters(fiber, system, beta3_s3_per_km=beta3_s3_per_km)
    signal_w = dbm_to_w(launch)
    eta = nli_result.eta_equivalent_w_inv2
    nli_w = eta * signal_w**3
    ase_w = np.full_like(signal_w, int(spans) * parameters["ase_per_span_w"])
    trx_snr_db = float(_object_value(system, "transceiver_snr_db", 18.0))
    trx_w = (
        signal_w / (10.0 ** (trx_snr_db / 10.0))
        if include_transceiver_noise
        else np.zeros_like(signal_w)
    )
    total_noise_w = ase_w + nli_w + trx_w
    snr_linear = signal_w / total_noise_w
    snr_db = 10.0 * np.log10(snr_linear)

    polarizations = int(_object_value(system, "polarizations", 2))
    cut_channel = plan[cut]
    gap = (
        float(_object_value(system, "shannon_gap_db_record_only", 3.0))
        if shannon_gap_db is None
        else float(shannon_gap_db)
    )
    per_channel_rate_gbps = _format_limited_information_rate_gbps(
        snr_linear,
        symbol_rate_hz=cut_channel.symbol_rate_hz,
        modulation=cut_channel.modulation,
        polarizations=polarizations,
        coding_rate=coding_rate,
        shannon_gap_db=gap,
    )
    return {
        "launch_dbm": launch,
        "spans": np.full(launch.shape, int(spans), dtype=int),
        "distance_km": np.full(launch.shape, int(spans) * parameters["span_length_km"]),
        "signal_w": signal_w,
        "ase_w": ase_w,
        "nli_w": nli_w,
        "trx_equivalent_noise_w": trx_w,
        "total_noise_w": total_noise_w,
        "snr_linear": snr_linear,
        "snr_db": snr_db,
        "format_limited_information_rate_per_channel_gbps": per_channel_rate_gbps,
        "format_limited_information_rate_total_tbps": per_channel_rate_gbps * len(plan) / 1000.0,
        "nli_integral": nli_result,
        "eta_equivalent_w_inv2": eta,
        "metadata": {
            "paper_reference": PAPER_REFERENCE,
            "power_sweep_rule": "All WDM channel powers scale together from the 0 dBm reference plan.",
            "modulation_note": (
                "Modulation limits the reported information-rate estimate; the GN NLI "
                "calculation remains Gaussian unless an independently justified correction is supplied."
            ),
            "coding_rate": coding_rate,
            "shannon_gap_db": gap,
            "fiber": _object_to_dict(fiber),
            "system": _object_to_dict(system),
            "options": asdict(options),
            "derived_parameters": parameters,
            "nli_metadata": dict(nli_result.metadata),
        },
    }


def make_launch_grid(
    launch_min_dbm: float = -10.0,
    launch_max_dbm: float = 10.0,
    launch_step_db: float = 0.25,
) -> np.ndarray:
    """Return an inclusive launch-power grid."""

    if launch_step_db <= 0.0 or launch_max_dbm <= launch_min_dbm:
        raise ValueError("launch_min_dbm < launch_max_dbm and launch_step_db > 0 are required.")
    count = int(np.floor((launch_max_dbm - launch_min_dbm) / launch_step_db))
    result = launch_min_dbm + np.arange(count + 1) * launch_step_db
    if result[-1] < launch_max_dbm - 1e-12:
        result = np.append(result, launch_max_dbm)
    return result


# ---------------------------------------------------------------------------
# Lightweight CLI and file helpers
# ---------------------------------------------------------------------------


def _write_curve_csv(result: Mapping[str, Any], path: str | Path) -> Path:
    """Save scalar sweep arrays only; integral metadata is saved alongside as JSON."""

    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    array_keys = [
        key
        for key, value in result.items()
        if isinstance(value, np.ndarray) and value.ndim == 1 and value.size == len(result["launch_dbm"])
    ]
    with destination.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(array_keys)
        writer.writerows(zip(*[result[key].tolist() for key in array_keys]))
    metadata_path = destination.with_suffix(".metadata.json")
    metadata = dict(result["metadata"])
    metadata["eta_equivalent_w_inv2"] = float(result["eta_equivalent_w_inv2"])
    metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2, default=_json_default), encoding="utf-8")
    return destination


def _json_default(value: Any) -> Any:
    if isinstance(value, np.generic):
        return value.item()
    if isinstance(value, np.ndarray):
        return value.tolist()
    if isinstance(value, GNIntegralResult):
        return value.as_dict()
    raise TypeError(f"Cannot serialize {type(value)!r} to JSON.")


def run_cli(argv: Optional[Sequence[str]] = None) -> Dict[str, Any]:
    parser = argparse.ArgumentParser(description="Numerical GN integral: Poggiolini Eq. (96)/(G.4)")
    parser.add_argument("--channels", type=int, default=90)
    parser.add_argument("--symbol-rate-gbd", type=float, default=95.0)
    parser.add_argument("--spacing-ghz", type=float, default=95.0)
    parser.add_argument("--span-length-km", type=float, default=80.0)
    parser.add_argument("--spans", type=int, default=30)
    parser.add_argument("--attenuation-db-per-km", type=float, default=0.166)
    parser.add_argument("--effective-area-um2", type=float, default=125.0)
    parser.add_argument("--dispersion-ps-nm-km", type=float, default=21.0)
    parser.add_argument("--n2-m2-w", type=float, default=2.2e-20)
    parser.add_argument("--beta3-s3-per-km", type=float, default=0.0)
    parser.add_argument("--noise-figure-db", type=float, default=5.0)
    parser.add_argument("--transceiver-snr-db", type=float, default=18.0)
    parser.add_argument("--modulation", default="64QAM")
    parser.add_argument("--roll-off", type=float, default=0.0)
    parser.add_argument("--launch-min", type=float, default=-10.0)
    parser.add_argument("--launch-max", type=float, default=10.0)
    parser.add_argument("--launch-step", type=float, default=0.25)
    parser.add_argument("--quadrature-order", type=int, default=24)
    parser.add_argument("--output-quadrature-order", type=int, default=3)
    parser.add_argument("--incoherent", action="store_true")
    parser.add_argument("--convergence", default=None, help="e.g. 8,12,16; reports eta convergence before the sweep")
    parser.add_argument("--save-csv", default=None)
    args = parser.parse_args(argv)

    fiber = FiberParameters(
        attenuation_db_per_km=args.attenuation_db_per_km,
        effective_area_um2=args.effective_area_um2,
        dispersion_ps_nm_km=args.dispersion_ps_nm_km,
        n2_m2_w=args.n2_m2_w,
        beta3_s3_per_km=args.beta3_s3_per_km,
    )
    system = SystemParameters(
        channels=args.channels,
        symbol_rate_gbd=args.symbol_rate_gbd,
        spacing_ghz=args.spacing_ghz,
        span_length_km=args.span_length_km,
        noise_figure_db=args.noise_figure_db,
        transceiver_snr_db=args.transceiver_snr_db,
    )
    options = GNIntegralOptions(
        quadrature_order=args.quadrature_order,
        output_quadrature_order=args.output_quadrature_order,
        coherent_accumulation=not args.incoherent,
    )
    plan = make_uniform_wdm(system, launch_power_dbm=0.0, roll_off=args.roll_off, modulation=args.modulation)
    if args.convergence:
        orders = [int(value.strip()) for value in args.convergence.split(",") if value.strip()]
        table = convergence_study(plan, quadrature_orders=orders, fiber=fiber, system=system, spans=args.spans, options=options)
        print("Quadrature convergence")
        for row in table:
            print(row)

    result = simulate_snr_curve(
        make_launch_grid(args.launch_min, args.launch_max, args.launch_step),
        args.spans,
        fiber=fiber,
        system=system,
        options=options,
        modulation=args.modulation,
        roll_off=args.roll_off,
    )
    integral = result["nli_integral"]
    peak = int(np.argmax(result["snr_db"]))
    print(f"eta_equivalent = {result['eta_equivalent_w_inv2']:.6e} W^-2")
    print(f"NLI integral runtime = {integral.metadata['runtime_s']:.2f} s")
    print(
        "SNR peak = "
        f"{result['snr_db'][peak]:.3f} dB at {result['launch_dbm'][peak]:.3f} dBm"
    )
    if args.save_csv:
        destination = _write_curve_csv(result, args.save_csv)
        print(f"Saved curve to: {destination}")
    return result


def main(argv: Optional[Sequence[str]] = None) -> Dict[str, Any]:
    """CLI-compatible entry point."""

    return run_cli(argv)


if __name__ == "__main__":
    main()
