#!/usr/bin/env python3
"""수치적분 GN model: closed-form NLI 식이나 fitting 계수를 사용하지 않습니다.

Based on the input/SNR structure of kimheeseo/LSCNS/2026_KICS_Fall/main.py
Git blob SHA: 7acd0ca1379592b6ad79acbe2748da4309830083 (retrieved 2026-09-09).
Standalone: Python >= 3.10, numpy >= 1.24, scipy >= 1.10; optional matplotlib.

Reference: Poggiolini et al., 'A Detailed Analytical Derivation of the GN
Model of Non-Linear Interference in Coherent Optical Transmission Systems',
https://arxiv.org/abs/1209.0394, Eq. (96)/(G.4), with beta3 = 0.

G_NLI(f) = (16/27)*gamma**2 * double_integral[
    Gtx(f1)*Gtx(f2)*Gtx(f1+f2-f)*rho(delta_beta)*A_N(delta_beta*L)] df1 df2
rho(db) = |(1-exp(-(alpha_power-1j*db)*L))/(alpha_power-1j*db)|**2
A_N(phi) = |sum(exp(1j*n*phi), n=0..N-1)|**2  [coherent, default]
           N                                  [incoherent, explicit option]
P_NLI = integral G_NLI(f)*|H_rx(f)|**2 df = eta_link * P_channel**3.

P_channel and Gtx are TOTAL DUAL-POLARIZATION quantities. The full integral
coefficient is 16/27, not the 8/27 coefficient of the original closed form.
The paper's field attenuation alpha is alpha_power/2 in this code.

We integrate the exact finite-span kernel with randomized Sobol quadrature
and change-of-variable importance sampling. No phase averaging, center-PSD
times bandwidth approximation, asinh closed-form eta, or empirical scaling.
The analytic z integral in rho is exact for a uniform span; the GN frequency
integrals are evaluated numerically. Scramble-based uncertainty estimates
measure NUMERICAL integration uncertainty, not experimental/model error.

Scope: identical spans, ideal loss-compensating lumped EDFAs, constant alpha/
beta2/gamma, equal-power/equal-rate uniformly spaced nonoverlapping channels,
rectangular or raised-cosine PSD and matched receiver. Not EGN or SSFM.
No ISRS/Raman, beta3, PMD/PDL, ROADM filtering, or signal-ASE nonlinearity.

Examples:
  python main_numerical_gn.py --spans 30 --save-dir gn_results --plot
  python main_numerical_gn.py --spans 10 30 50 --rtol 0.001 --max-power 23
  python main_numerical_gn.py --accumulation incoherent --save-dir gn_i
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
import time
import warnings
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Callable, Mapping, Sequence

import numpy as np
import scipy
from scipy.stats import qmc, t as student_t

C0 = 299_792_458.0
H_PLANCK = 6.626_070_15e-34
MANAKOV_DP = 16.0 / 27.0
SOURCE_URL = "https://github.com/kimheeseo/LSCNS/blob/main/2026_KICS_Fall/main.py"
SOURCE_SHA = "7acd0ca1379592b6ad79acbe2748da4309830083"
REFERENCE_URL = "https://arxiv.org/abs/1209.0394"


@dataclass(frozen=True)
class FiberParameters:
    name: str = "G.654.E"
    wavelength_nm: float = 1550.0
    attenuation_db_per_km: float = 0.166
    effective_area_um2: float = 125.0
    dispersion_ps_nm_km: float = 21.0
    n2_m2_w: float = 2.2e-20
    alpha_power_per_km: float | None = None
    beta2_s2_per_km: float | None = None
    gamma_per_w_km: float | None = None


@dataclass(frozen=True)
class SystemParameters:
    channels: int = 90
    symbol_rate_gbd: float = 95.0
    spacing_ghz: float = 95.0
    span_length_km: float = 80.0
    noise_figure_db: float = 5.0
    transceiver_snr_db: float = 18.0
    polarizations: int = 2
    ase_bandwidth_hz: float | None = None
    stated_gain_bandwidth_thz: float | None = 4.8
    qam_order_record_only: int = 64
    shannon_gap_db_record_only: float = 3.0


@dataclass(frozen=True)
class NumericalGNOptions:
    accumulation: str = "coherent"
    rolloff: float = 0.0
    cut_index: int | None = None  # zero-based; default channels//2 is a real channel
    min_power: int = 17          # initial points PER scramble = 2**min_power
    max_power: int = 21         # doubles until convergence or this limit
    replicates: int = 8         # independent scrambles, not repeated same points
    relative_tolerance: float = 0.005
    seed: int = 20260909
    chunk_power: int = 15       # limits temporary array memory, not precision


def _positive(name: str, value: float, allow_zero: bool = False) -> None:
    if not np.isfinite(value) or (value < 0 if allow_zero else value <= 0):
        raise ValueError(f"{name}: finite {'nonnegative' if allow_zero else 'positive'} value required")


def _integer(name: str, value: int, minimum: int = 1) -> int:
    if isinstance(value, (bool, np.bool_)) or not isinstance(value, (int, np.integer)) or value < minimum:
        raise ValueError(f"{name}: integer >= {minimum} required")
    return int(value)


def derive_parameters(fiber: FiberParameters = FiberParameters(),
                      system: SystemParameters = SystemParameters()) -> dict:
    """원본과 같은 단위 변환, loss compensation, ASE convention."""
    _integer("channels", system.channels)
    if system.polarizations != 2:
        raise ValueError("This implementation requires total dual-polarization channel power.")
    for key in ("wavelength_nm", "effective_area_um2", "n2_m2_w"):
        _positive(key, getattr(fiber, key))
    for key in ("symbol_rate_gbd", "spacing_ghz", "span_length_km"):
        _positive(key, getattr(system, key))
    _positive("noise_figure_db", system.noise_figure_db, allow_zero=True)
    if not np.isfinite(system.transceiver_snr_db):
        raise ValueError("transceiver_snr_db must be finite; disable TRX explicitly instead.")
    if system.ase_bandwidth_hz is not None:
        _positive("ase_bandwidth_hz", system.ase_bandwidth_hz)
    if system.stated_gain_bandwidth_thz is not None:
        _positive("stated_gain_bandwidth_thz", system.stated_gain_bandwidth_thz)
    wavelength = fiber.wavelength_nm * 1e-9
    rs = system.symbol_rate_gbd * 1e9
    alpha = (fiber.attenuation_db_per_km * np.log(10) / 10
             if fiber.alpha_power_per_km is None else fiber.alpha_power_per_km)
    beta2 = (-wavelength**2 / (2 * np.pi * C0) * fiber.dispersion_ps_nm_km * 1e-3
             if fiber.beta2_s2_per_km is None else fiber.beta2_s2_per_km)
    gamma = (2 * np.pi * fiber.n2_m2_w / (wavelength * fiber.effective_area_um2 * 1e-12) * 1e3
             if fiber.gamma_per_w_km is None else fiber.gamma_per_w_km)
    _positive("alpha_power_per_km", alpha)
    _positive("gamma_per_w_km", gamma, allow_zero=True)
    if not np.isfinite(beta2):
        raise ValueError("beta2 must be finite")
    optical_depth = alpha * system.span_length_km
    if optical_depth > 100 or abs(system.transceiver_snr_db) > 300 or system.noise_figure_db > 100:
        raise ValueError("Gain/noise inputs outside supported numerical range")
    gain = np.exp(optical_depth)
    ase_bw = rs if system.ase_bandwidth_hz is None else system.ase_bandwidth_hz
    return {
        "wavelength_m": wavelength, "frequency_hz": C0 / wavelength,
        "rate_hz": rs, "spacing_hz": system.spacing_ghz * 1e9,
        "alpha_power_per_km": float(alpha), "alpha_field_per_km": float(alpha / 2),
        "beta2_s2_per_km": float(beta2), "beta2_ps2_per_km": float(beta2 * 1e24),
        "gamma_per_w_km": float(gamma),
        "effective_length_km": float(-np.expm1(-optical_depth) / alpha),
        "asymptotic_length_km": float(1 / alpha),
        "attenuation_db_per_km_used": float(alpha * 10 / np.log(10)),
        "span_gain_db": float(optical_depth * 10 / np.log(10)),
        "span_gain_linear": float(gain), "ase_bandwidth_hz": float(ase_bw),
        "ase_per_span_w": float(H_PLANCK * C0 / wavelength * 10**(system.noise_figure_db / 10)
                                * ase_bw * np.expm1(optical_depth)),
        "wdm_occupied_bandwidth_thz": system.channels * system.spacing_ghz / 1000,
    }


def _validate_options(options: NumericalGNOptions, system: SystemParameters) -> int:
    if options.accumulation not in ("coherent", "incoherent"):
        raise ValueError("accumulation must be coherent or incoherent")
    if not np.isfinite(options.rolloff) or not 0 <= options.rolloff <= 1:
        raise ValueError("rolloff must lie in [0, 1]")
    if system.channels > 1 and system.spacing_ghz < system.symbol_rate_gbd * (1 + options.rolloff) * (1 - 1e-12):
        raise ValueError("Overlapping WDM channels are not supported: spacing >= Rs*(1+rolloff) required.")
    cut = system.channels // 2 if options.cut_index is None else _integer("cut_index", options.cut_index, 0)
    if cut >= system.channels:
        raise ValueError("cut_index must be less than channels")
    for key in ("min_power", "max_power", "chunk_power"):
        _integer(key, getattr(options, key))
    if not 4 <= options.min_power < options.max_power <= 28:
        raise ValueError("Require 4 <= min_power < max_power <= 28 (at least two refinement levels).")
    _integer("replicates", options.replicates, 4)
    _integer("seed", options.seed, 0)
    if not 0 < options.relative_tolerance < 1:
        raise ValueError("relative_tolerance must lie strictly between 0 and 1")
    return cut


def raised_cosine_psd(x: np.ndarray, rolloff: float = 0.0) -> np.ndarray:
    """Dimensionless RC power spectrum; integral over x=f/Rs is one."""
    ax = np.abs(np.asarray(x, dtype=float))
    if rolloff == 0:
        return (ax <= 0.5).astype(float)
    transition = np.clip((ax - (1 - rolloff) / 2) / rolloff, 0, 1)
    return 0.5 * (1 + np.cos(np.pi * transition))


def wdm_psd_shape(x: np.ndarray, system: SystemParameters,
                  options: NumericalGNOptions) -> np.ndarray:
    """S(x), such that Gtx(f)=P_channel/Rs*S(f/Rs). Gaps remain zero."""
    x = np.asarray(x, dtype=float)
    if system.channels == 1:
        return raised_cosine_psd(x, options.rolloff)
    cut = system.channels // 2 if options.cut_index is None else options.cut_index
    spacing = system.spacing_ghz / system.symbol_rate_gbd
    first = -cut * spacing
    nearest = np.rint((x - first) / spacing).astype(np.int64)
    valid = (nearest >= 0) & (nearest < system.channels)
    offset = x - (first + nearest * spacing)
    return np.where(valid, raised_cosine_psd(offset, options.rolloff), 0.0)


def coherent_accumulation(phase: np.ndarray, spans: int) -> np.ndarray:
    """Exact finite geometric sum squared, stable at phase = 2*pi*k."""
    spans = _integer("spans", spans)
    phase = np.asarray(phase, dtype=float)
    if spans == 1:
        return np.ones_like(phase)
    wrapped = np.remainder(phase + np.pi, 2 * np.pi) / (2 * np.pi) - 0.5
    return (spans * np.sinc(spans * wrapped) / np.sinc(wrapped))**2


def single_span_kernel(delta_beta: np.ndarray, alpha_power: float, length: float) -> np.ndarray:
    """rho in km^2. Finite-length numerator retained even at high loss."""
    db = np.asarray(delta_beta, dtype=float)
    loss = np.exp(-alpha_power * length)
    numerator = np.expm1(-alpha_power * length)**2 + 4 * loss * np.sin(db * length / 2)**2
    return numerator / (alpha_power**2 + db**2)


def _sample_integrand(points: np.ndarray, parameters: Mapping, system: SystemParameters,
                      options: NumericalGNOptions) -> tuple[np.ndarray, np.ndarray]:
    """Return importance-weighted single-span eta samples and EXACT phases.

    x = f/Rs; u = (f1-f)/Rs; v = (f2-f)/Rs. All powers of Rs cancel
    between the three PSD factors and df1*df2*df. Sinh in u resolves the
    narrow phase-matching ridges; arctan in v cancels the Lorentz denominator.
    These are changes of variables with their Jacobians, not fitted models.
    """
    alpha = parameters["alpha_power_per_km"]
    kappa = 4 * np.pi**2 * parameters["beta2_s2_per_km"] * parameters["rate_hz"]**2
    gamma = parameters["gamma_per_w_km"]
    length = system.span_length_km
    cut = system.channels // 2 if options.cut_index is None else options.cut_index
    spacing = system.spacing_ghz / system.symbol_rate_gbd
    half = (1 + options.rolloff) / 2
    band_lo, band_hi = -cut * spacing - half, (system.channels - 1 - cut) * spacing + half
    x = (2 * points[:, 0] - 1) * half
    u_lo, u_hi = band_lo - x, band_hi - x
    # For very weak dispersion use uniform variables to avoid small-number loss.
    weak = abs(kappa) * (band_hi - band_lo + 2 * half)**2 / alpha < 1e-4
    if weak:
        u_jac = u_hi - u_lo
        u = u_lo + points[:, 1] * u_jac
    else:
        scale = alpha / (abs(kappa) * (band_hi - band_lo + 2 * half))
        t_lo, t_hi = np.arcsinh(u_lo / scale), np.arcsinh(u_hi / scale)
        t = t_lo + points[:, 1] * (t_hi - t_lo)
        u = scale * np.sinh(t)
        u_jac = scale * np.cosh(t) * (t_hi - t_lo)
    v_lo = np.maximum(band_lo - x, band_lo - x - u)
    v_hi = np.minimum(band_hi - x, band_hi - x - u)
    b = np.abs(kappa * u)
    ratio = b / alpha
    near_zero = ratio * np.maximum(np.abs(v_lo), np.abs(v_hi)) < 1e-5
    v = np.empty_like(u)
    kernel_jac = np.empty_like(u)
    # Exact uniform transform where the arctan transform would be ill-conditioned.
    v[near_zero] = v_lo[near_zero] + points[near_zero, 2] * (v_hi - v_lo)[near_zero]
    kernel_jac[near_zero] = (v_hi - v_lo)[near_zero] / (
        alpha**2 + (kappa * u[near_zero] * v[near_zero])**2)
    far = ~near_zero
    q = ratio[far]
    theta_lo = np.arctan(q * v_lo[far])
    # atan2 avoids cancellation when both atan endpoints approach +/-pi/2.
    theta_width = np.arctan2(q * (v_hi - v_lo)[far], 1 + (q * v_lo[far]) * (q * v_hi[far]))
    theta = theta_lo + points[far, 2] * theta_width
    v[far] = np.tan(theta) / q
    kernel_jac[far] = theta_width / (alpha * b[far])
    phase = kappa * u * v * length
    numerator = np.expm1(-alpha * length)**2 + 4 * np.exp(-alpha * length) * np.sin(phase / 2)**2
    spectra = (wdm_psd_shape(x + u, system, options) * wdm_psd_shape(x + v, system, options)
               * wdm_psd_shape(x + u + v, system, options))
    weight = (MANAKOV_DP * gamma**2 * (2 * half) * raised_cosine_psd(x, options.rolloff)
              * spectra * u_jac * numerator * kernel_jac)
    if not np.all(np.isfinite(weight)) or not np.all(np.isfinite(phase)):
        raise FloatingPointError("Non-finite integrand. Check physical inputs/range.")
    return weight, phase


def _relative(numerator: np.ndarray, denominator: np.ndarray) -> np.ndarray:
    return np.divide(numerator, denominator, out=np.zeros_like(numerator), where=denominator != 0)


def numerical_gn_eta(spans: int | Sequence[int] = 30, *,
                     fiber: FiberParameters = FiberParameters(),
                     system: SystemParameters = SystemParameters(),
                     options: NumericalGNOptions = NumericalGNOptions(),
                     progress: Callable[[dict], None] | None = None) -> dict:
    """Compute eta_link [W^-2] for one/multiple span counts, with convergence.

    Refines nested 2**m Sobol point sets. Independent scrambles estimate
    standard error; Student-t produces an estimated 95% half-width.
    Stop only when BOTH relative half-width and change from the preceding
    level meet the tolerance for ALL requested span counts. This is a
    diagnostic, not a rigorous error bound; check another seed for new cases.
    Unconverged estimates are returned with converged=False and a warning.
    """
    parameters = derive_parameters(fiber, system)
    cut = _validate_options(options, system)
    raw_spans = [spans] if np.isscalar(spans) else list(spans)
    counts = sorted(set(_integer("spans", n) for n in raw_spans))
    if not counts:
        raise ValueError("spans cannot be empty")
    messages = []
    support_thz = ((system.channels - 1) * system.spacing_ghz
                   + system.symbol_rate_gbd * (1 + options.rolloff)) / 1000
    if system.stated_gain_bandwidth_thz is not None and support_thz > system.stated_gain_bandwidth_thz:
        messages.append(f"WDM support {support_thz:g} THz exceeds stated amplifier bandwidth "
                        f"{system.stated_gain_bandwidth_thz:g} THz. Inputs retained; ideal flat gain assumed.")
    if support_thz > 5:
        messages.append("Wideband case: ISRS and frequency-dependent fiber/amplifier parameters are omitted.")
    if system.ase_bandwidth_hz is not None and not np.isclose(system.ase_bandwidth_hz, parameters["rate_hz"]):
        messages.append("ASE bandwidth override differs from matched-filter noise bandwidth Rs; legacy ASE convention retained.")
    if parameters["beta2_s2_per_km"] == 0:
        messages.append("Zero dispersion is allowed for mathematical tests; GN Gaussianity is not physically validated here.")
    for message in messages:
        warnings.warn(message, RuntimeWarning, stacklevel=2)
    engines = [qmc.Sobol(d=3, scramble=True, seed=int(s.generate_state(1)[0]))
               for s in np.random.SeedSequence(options.seed).spawn(options.replicates)]
    sums = np.zeros((options.replicates, len(counts)))
    history, previous, total = [], None, 0
    critical = float(student_t.ppf(0.975, options.replicates - 1))
    started = time.perf_counter()
    converged = False
    for power in range(options.min_power, options.max_power + 1):
        additional = 2**power - total
        for rep, engine in enumerate(engines):
            remaining = additional
            while remaining:
                size = min(remaining, 2**options.chunk_power)
                # Successive contiguous dyadic blocks; each tested total is 2**power.
                points = engine.random(size)
                base, phase = _sample_integrand(points, parameters, system, options)
                for j, n in enumerate(counts):
                    if options.accumulation == "incoherent":
                        sums[rep, j] += n * np.sum(base, dtype=np.float64)
                    else:
                        sums[rep, j] += np.sum(base * coherent_accumulation(phase, n), dtype=np.float64)
                remaining -= size
        total = 2**power
        estimates = sums / total
        mean = np.mean(estimates, axis=0)
        half95 = critical * np.std(estimates, axis=0, ddof=1) / np.sqrt(options.replicates)
        rel95 = _relative(half95, mean)
        change = None if previous is None else _relative(np.abs(mean - previous), mean)
        converged = bool(change is not None and np.all(rel95 <= options.relative_tolerance)
                         and np.all(change <= options.relative_tolerance))
        state = {"power": power, "samples_per_replicate": total,
                 "eta_link_w_inv2": mean.tolist(), "relative_ci95_halfwidth": rel95.tolist(),
                 "relative_change": None if change is None else change.tolist(),
                 "elapsed_seconds": time.perf_counter() - started, "converged": converged}
        history.append(state)
        if progress:
            progress(state)
        if converged:
            break
        previous = mean.copy()
    if not converged:
        message = "Integration did NOT meet the requested tolerance; increase max_power/replicates and verify with another seed."
        messages.append(message)
        warnings.warn(message, RuntimeWarning, stacklevel=2)
    return {
        "spans": counts, "eta_link_w_inv2": mean.tolist(),
        "ci95_halfwidth_w_inv2": half95.tolist(), "relative_ci95_halfwidth": rel95.tolist(),
        "replicate_eta_link_w_inv2": estimates.tolist(), "converged": converged,
        "history": history, "samples_per_replicate": total,
        "total_integrand_evaluations": total * options.replicates,
        "elapsed_seconds": time.perf_counter() - started,
        "fiber": asdict(fiber), "system": asdict(system), "options": asdict(options),
        "parameters": parameters, "cut_index_used": cut, "wdm_support_thz": support_thz,
        "power_definition": "total_dp_per_channel", "integral_coefficient": MANAKOV_DP,
        "uncertainty_kind": "estimated numerical integration uncertainty; NOT physical/model validation error",
        "warnings": messages, "source_url": SOURCE_URL, "source_git_blob_sha": SOURCE_SHA,
        "reference_url": REFERENCE_URL,
        "versions": {"python": sys.version.split()[0], "numpy": np.__version__, "scipy": scipy.__version__},
    }


def calculate_snr(launch_dbm: Sequence[float] | float, spans: int, eta_link: float,
                  parameters: Mapping, system: SystemParameters, *,
                  include_transceiver_noise: bool = True) -> dict[str, np.ndarray]:
    """Same SNR budget as main.py; eta_link ALREADY includes all spans.

    IMPORTANT migration difference: third argument is eta_link, NOT eta_per_span.
    Do not multiply NLI by spans again. P is total DP power PER CHANNEL.
    """
    spans = _integer("spans", spans)
    _positive("eta_link", eta_link, allow_zero=True)
    launch = np.atleast_1d(np.asarray(launch_dbm, dtype=float))
    if launch.ndim != 1 or launch.size == 0 or not np.all(np.isfinite(launch)) or np.any(np.abs(launch) > 100):
        raise ValueError("launch_dbm must be a nonempty finite 1-D array within [-100, 100]")
    signal = 1e-3 * 10**(launch / 10)
    ase = np.full_like(signal, spans * parameters["ase_per_span_w"])
    nli = eta_link * signal**3
    trx = signal / 10**(system.transceiver_snr_db / 10) if include_transceiver_noise else np.zeros_like(signal)
    noise = ase + nli + trx
    snr = signal / noise
    return {"launch_dbm": launch, "spans": np.full(launch.shape, spans, dtype=int),
            "distance_km": np.full(launch.shape, spans * system.span_length_km),
            "signal_w": signal, "ase_w": ase, "nli_w": nli,
            "trx_equivalent_noise_w": trx, "total_noise_w": noise,
            "snr_linear": snr, "snr_db": 10 * np.log10(snr),
            "eta_link_w_inv2": np.full(launch.shape, eta_link)}


def simulate_snr_curve(launch_dbm: Sequence[float] | float, spans: int,
                       fiber: FiberParameters = FiberParameters(),
                       system: SystemParameters = SystemParameters(),
                       options: NumericalGNOptions = NumericalGNOptions(), *,
                       include_transceiver_noise: bool = True) -> dict:
    """Single-span-count convenience API. Returns columns, like main.py without pandas.

    For sweeps over span counts, use numerical_gn_eta once, then calculate_snr
    for each eta_link. Do not redo the integral for every launch power.
    """
    result = numerical_gn_eta(spans, fiber=fiber, system=system, options=options)
    if not result["converged"]:
        raise RuntimeError("Unconverged integral. Inspect numerical_gn_eta diagnostics or raise max_power.")
    return calculate_snr(launch_dbm, spans, result["eta_link_w_inv2"][0], result["parameters"], system,
                         include_transceiver_noise=include_transceiver_noise)


def analytical_optimum_launch_dbm(eta_link: float, spans: int, parameters: Mapping) -> float:
    """Exact optimum of this scalar noise budget, NOT a closed-form GN eta.

    Popt=(N*ASE_per_span/(2*eta_link))**(1/3). Constant TRX SNR does not
    shift this optimum. This does not impose a reach/BER or FEC threshold.
    """
    _positive("eta_link", eta_link)
    _integer("spans", spans)
    power = (spans * parameters["ase_per_span_w"] / (2 * eta_link))**(1 / 3)
    return float(10 * np.log10(power / 1e-3))


def error_metrics(reference_snr_db: Sequence[float], predicted_snr_db: Sequence[float]) -> dict:
    """External-reference metrics on the SAME points; never used to fit eta."""
    ref, pred = np.asarray(reference_snr_db, float), np.asarray(predicted_snr_db, float)
    if ref.shape != pred.shape or ref.ndim != 1 or ref.size == 0 or not np.all(np.isfinite([ref, pred])):
        raise ValueError("reference and prediction must be finite equal-length 1-D arrays")
    residual = pred - ref
    return {"samples": int(ref.size), "bias_db": float(np.mean(residual)),
            "mae_db": float(np.mean(np.abs(residual))),
            "rmse_db": float(np.sqrt(np.mean(residual**2))),
            "max_abs_db": float(np.max(np.abs(residual))),
            "mape_linear_snr_percent": float(np.mean(np.abs(10**(residual / 10) - 1)) * 100)}


def _progress(state: dict) -> None:
    rel = max(state["relative_ci95_halfwidth"]) * 100
    change = state["relative_change"]
    change_text = "n/a" if change is None else f"{max(change)*100:.4f}%"
    print(f"2^{state['power']}/scramble: estimated 95% half-width <= {rel:.4f}%, "
          f"level change <= {change_text}, {state['elapsed_seconds']:.1f}s", flush=True)


def run_cli(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--spans", type=int, nargs="+", default=[30])
    parser.add_argument("--channels", type=int, default=90)
    parser.add_argument("--symbol-rate-gbd", type=float, default=95)
    parser.add_argument("--spacing-ghz", type=float, default=95)
    parser.add_argument("--span-length-km", type=float, default=80)
    parser.add_argument("--attenuation-db-per-km", type=float, default=0.166)
    parser.add_argument("--dispersion-ps-nm-km", type=float, default=21)
    parser.add_argument("--effective-area-um2", type=float, default=125)
    parser.add_argument("--n2-m2-w", type=float, default=2.2e-20)
    parser.add_argument("--wavelength-nm", type=float, default=1550)
    parser.add_argument("--noise-figure-db", type=float, default=5)
    parser.add_argument("--transceiver-snr-db", type=float, default=18)
    parser.add_argument("--stated-gain-bandwidth-thz", type=float, default=4.8,
                        help="Informational only; never used to truncate channels")
    parser.add_argument("--no-trx", action="store_true")
    parser.add_argument("--accumulation", choices=["coherent", "incoherent"], default="coherent")
    parser.add_argument("--rolloff", type=float, default=0)
    parser.add_argument("--cut-index", type=int)
    parser.add_argument("--min-power", type=int, default=17)
    parser.add_argument("--max-power", type=int, default=21)
    parser.add_argument("--replicates", type=int, default=8)
    parser.add_argument("--rtol", type=float, default=0.005)
    parser.add_argument("--seed", type=int, default=20260909)
    parser.add_argument("--launch-min-dbm", type=float, default=-10)
    parser.add_argument("--launch-max-dbm", type=float, default=10)
    parser.add_argument("--launch-step-dbm", type=float, default=0.25)
    parser.add_argument("--save-dir", type=Path)
    parser.add_argument("--overwrite", action="store_true", help="Allow replacing this program's named output files")
    parser.add_argument("--allow-unconverged", action="store_true", help="Keep warning/flag but exit 0 instead of 2")
    parser.add_argument("--plot", action="store_true")
    parser.add_argument("--reference-csv", type=Path,
                        help="CSV columns: spans, launch_dbm, snr_db; same conditions required, no fitting")
    args = parser.parse_args(argv)
    if not np.all(np.isfinite([args.launch_min_dbm, args.launch_max_dbm, args.launch_step_dbm])):
        parser.error("Launch grid must be finite")
    if args.launch_step_dbm <= 0 or not -100 <= args.launch_min_dbm <= args.launch_max_dbm <= 100:
        parser.error("Invalid launch range or step")
    intervals = int(np.floor((args.launch_max_dbm - args.launch_min_dbm) / args.launch_step_dbm + 1e-10))
    if intervals > 1_000_000:
        parser.error("Launch grid too large")
    launch = args.launch_min_dbm + np.arange(intervals + 1) * args.launch_step_dbm
    if args.plot and args.save_dir is None:
        parser.error("--plot requires --save-dir")
    if args.save_dir:
        targets = ["gn_report.json", "snr_curves.csv"] + (["snr_curves.png"] if args.plot else [])
        existing = [str(args.save_dir / name) for name in targets if (args.save_dir / name).exists()]
        if existing and not args.overwrite:
            parser.error("Output exists; select a new --save-dir or explicitly --overwrite: " + ", ".join(existing))
    fiber = FiberParameters(wavelength_nm=args.wavelength_nm, attenuation_db_per_km=args.attenuation_db_per_km,
                            dispersion_ps_nm_km=args.dispersion_ps_nm_km, effective_area_um2=args.effective_area_um2,
                            n2_m2_w=args.n2_m2_w)
    system = SystemParameters(channels=args.channels, symbol_rate_gbd=args.symbol_rate_gbd,
                              spacing_ghz=args.spacing_ghz, span_length_km=args.span_length_km,
                              noise_figure_db=args.noise_figure_db, transceiver_snr_db=args.transceiver_snr_db,
                              stated_gain_bandwidth_thz=args.stated_gain_bandwidth_thz)
    options = NumericalGNOptions(accumulation=args.accumulation, rolloff=args.rolloff, cut_index=args.cut_index,
                                 min_power=args.min_power, max_power=args.max_power, replicates=args.replicates,
                                 relative_tolerance=args.rtol, seed=args.seed)
    result = numerical_gn_eta(args.spans, fiber=fiber, system=system, options=options, progress=_progress)
    result["include_transceiver_noise"] = not args.no_trx
    curves, optima = [], []
    for n, eta in zip(result["spans"], result["eta_link_w_inv2"]):
        curve = calculate_snr(launch, n, eta, result["parameters"], system, include_transceiver_noise=not args.no_trx)
        curves.append(curve)
        if eta > 0:
            optimum = analytical_optimum_launch_dbm(eta, n, result["parameters"])
            snr_opt = calculate_snr(optimum, n, eta, result["parameters"], system,
                                    include_transceiver_noise=not args.no_trx)["snr_db"][0]
            optima.append({"spans": n, "optimum_launch_dbm": optimum, "optimum_snr_db": float(snr_opt)})
            print(f"N={n}: eta_link={eta:.8g} W^-2, Popt={optimum:.4f} dBm/channel (total DP), GSNR={snr_opt:.4f} dB")
    result["optima"] = optima
    if args.reference_csv:
        with args.reference_csv.open(newline="", encoding="utf-8-sig") as stream:
            rows = list(csv.DictReader(stream))
        if not rows or not {"spans", "launch_dbm", "snr_db"}.issubset(rows[0]):
            parser.error("Reference CSV requires nonempty spans,launch_dbm,snr_db columns")
        lookup = dict(zip(result["spans"], result["eta_link_w_inv2"]))
        references, predictions = [], []
        for row in rows:
            n = int(row["spans"])
            if n not in lookup:
                parser.error(f"Reference span {n} was not requested in --spans")
            pred = calculate_snr(float(row["launch_dbm"]), n, lookup[n], result["parameters"], system,
                                 include_transceiver_noise=not args.no_trx)["snr_db"][0]
            references.append(float(row["snr_db"]))
            predictions.append(float(pred))
        result["reference_metrics"] = error_metrics(references, predictions)
        result["reference_csv"] = str(args.reference_csv)
        print("External-reference comparison (conditions must match):", result["reference_metrics"])
    if args.save_dir:
        args.save_dir.mkdir(parents=True, exist_ok=True)
        with (args.save_dir / "gn_report.json").open("w", encoding="utf-8") as stream:
            json.dump(result, stream, ensure_ascii=False, indent=2, allow_nan=False)
        with (args.save_dir / "snr_curves.csv").open("w", newline="", encoding="utf-8") as stream:
            writer = csv.writer(stream)
            writer.writerow([*curves[0], "integration_converged"])
            for curve in curves:
                writer.writerows([*row, result["converged"]] for row in zip(*curve.values()))
        if args.plot:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            fig, ax = plt.subplots(figsize=(8, 5))
            for curve in curves:
                ax.plot(curve["launch_dbm"], curve["snr_db"], label=f"{curve['spans'][0]} spans")
            ax.set(xlabel="Launch power per channel, total DP (dBm)", ylabel="GSNR (dB)",
                   title="Numerical GN: " + options.accumulation + ("" if result["converged"] else " [UNCONVERGED]"))
            ax.grid(alpha=0.3)
            ax.legend()
            fig.tight_layout()
            fig.savefig(args.save_dir / "snr_curves.png", dpi=180)
            plt.close(fig)
    print("Convergence:", result["converged"], "(numerical criterion only; no paper/experiment accuracy claim)")
    return 0 if result["converged"] or args.allow_unconverged else 2


if __name__ == "__main__":
    raise SystemExit(run_cli())
