"""Capacity optimization under optical and submarine power-feed constraints."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np

from .help_functions import dBm2Lin
from .snr_calculator_silas import SNR_GD


@dataclass(frozen=True)
class CapacityOptimizationResult:
    """Detailed result for all candidate span lengths."""

    eta_at_optimum: np.ndarray
    optimal_span_length_km: float
    maximum_capacity_bps: float
    gsnr_at_capacity_optimum_linear: float
    gsnr_by_span_linear: np.ndarray
    launch_power_by_span_w: np.ndarray
    repeater_count: np.ndarray
    span_count: np.ndarray
    unconstrained_nonlinear_threshold_w: np.ndarray
    feasible: np.ndarray


def _validate_monotonic(name: str, values: np.ndarray) -> np.ndarray:
    arr = np.asarray(values, dtype=float)
    if arr.ndim != 1 or len(arr) == 0:
        raise ValueError(f"{name} must be a non-empty one-dimensional array.")
    if np.any(~np.isfinite(arr)) or np.any(arr <= 0):
        raise ValueError(f"{name} must contain finite, strictly positive values.")
    if np.any(np.diff(arr) <= 0):
        raise ValueError(f"{name} must be strictly increasing.")
    return arr


def _extract_efficiency_grid(
    eta: np.ndarray, n_spans: int, n_powers: int
) -> np.ndarray:
    """Accept either a raw grid or the legacy grid with one header row/column."""
    grid = np.asarray(eta, dtype=float)
    if grid.shape == (n_spans, n_powers):
        result = grid
    elif grid.shape == (n_spans + 1, n_powers + 1):
        result = grid[1:, 1:]
    else:
        raise ValueError(
            "eta must have shape (len(L_span), len(Ps)) or the legacy "
            "headered shape (len(L_span)+1, len(Ps)+1)."
        )
    if np.any(~np.isfinite(result)) or np.any((result <= 0) | (result > 1)):
        raise ValueError("Electrical-to-optical efficiencies must be in (0, 1].")
    return result


def _link_counts(
    total_length_km: float,
    requested_span_km: np.ndarray,
    mode: Literal["physical", "paper"],
) -> tuple[np.ndarray, np.ndarray]:
    if mode == "physical":
        span_count = np.ceil(total_length_km / requested_span_km).astype(int)
        repeater_count = np.maximum(span_count - 1, 1)
    elif mode == "paper":
        # Reproduces the convention stated in the paper and legacy code.
        repeater_count = np.maximum(
            np.floor(total_length_km / requested_span_km).astype(int), 1
        )
        span_count = repeater_count.copy()
    else:
        raise ValueError("count_mode must be 'physical' or 'paper'.")
    return span_count, repeater_count


def optimized_capacity_extended(
    L_span,
    Ltot,
    NF,
    V_PFE,
    R,
    eta,
    epsilon,
    M,
    gamma,
    gamma_XT,
    alpha,
    l_FIFO,
    D,
    Rs,
    delta_f,
    Ps,
    lamb,
    B,
    L_span_min,
    *,
    n_cores: int = 1,
    count_mode: Literal["physical", "paper"] = "physical",
    gamma_xt_unit: str = "linear_per_km",
    nonlinear_backoff_db: float | None = 2.0,
    launch_power_cap_dbm: float | None = 18.0,
    implementation_gap_linear: float = 1.0,
    overhead_fraction: float = 0.0,
    return_details: bool = False,
):
    """Optimize cable capacity over candidate repeater span lengths.

    The original positional signature is retained. New keyword arguments add the
    paper's 2-dB nonlinear-threshold backoff and 18-dBm launch-power limit,
    realistic core counting, optional implementation gap/overhead, and an exact
    result object.
    """
    span_candidates = _validate_monotonic("L_span", L_span)
    launch_powers_w = _validate_monotonic("Ps", Ps)
    if Ltot <= 0 or NF < 0 or V_PFE <= 0 or R <= 0 or not 0 <= epsilon < 1:
        raise ValueError("Invalid link, amplifier or PFE parameters.")
    if not isinstance(M, (int, np.integer)) or M < 1:
        raise ValueError("M must be a positive integer number of fibre pairs.")
    if not isinstance(n_cores, (int, np.integer)) or n_cores < 1:
        raise ValueError("n_cores must be a positive integer.")
    if L_span_min <= 0:
        raise ValueError("L_span_min must be positive.")
    if implementation_gap_linear < 1:
        raise ValueError("implementation_gap_linear must be at least 1.")
    if not 0 <= overhead_fraction < 1:
        raise ValueError("overhead_fraction must be in [0, 1).")

    efficiency = _extract_efficiency_grid(
        eta, len(span_candidates), len(launch_powers_w)
    )
    span_count, repeater_count = _link_counts(Ltot, span_candidates, count_mode)
    n_channels = int(np.floor(B / delta_f + 1e-12))
    if n_channels < 1:
        raise ValueError("The bandwidth contains no complete WDM channel.")

    # Each fibre pair has two directions; each core requires its own amplifier.
    amplifiers_per_repeater = 2 * M * n_cores

    max_snr_linear = np.full(len(span_candidates), np.nan)
    max_power_w = np.full(len(span_candidates), np.nan)
    eta_at_optimum = np.full(len(span_candidates), np.nan)
    unconstrained_threshold_w = np.full(len(span_candidates), np.nan)
    feasible_span = np.zeros(len(span_candidates), dtype=bool)

    hardware_cap_w = (
        np.inf
        if launch_power_cap_dbm is None
        else float(dBm2Lin(launch_power_cap_dbm))
    )

    for i, _ in enumerate(span_candidates):
        snr_curve = SNR_GD(
            launch_powers_w,
            lamb,
            alpha,
            NF,
            int(span_count[i]),
            gamma,
            Ltot,
            B,
            D,
            Rs,
            gamma_XT,
            l_FIFO,
            delta_f,
            gamma_xt_unit=gamma_xt_unit,
        )
        nonlinear_idx = int(np.nanargmax(snr_curve))
        nonlinear_threshold_w = launch_powers_w[nonlinear_idx]
        unconstrained_threshold_w[i] = nonlinear_threshold_w
        nonlinear_cap_w = (
            np.inf
            if nonlinear_backoff_db is None
            else nonlinear_threshold_w * 10.0 ** (-nonlinear_backoff_db / 10.0)
        )

        # Available optical power per amplifier. eta is allowed to vary with
        # output power and span loss, so feasibility is evaluated point by point.
        total_available_optical_w = (
            V_PFE**2
            * (1.0 - epsilon)
            * efficiency[i, :]
            / (4.0 * R * Ltot)
        )
        pfe_cap_per_amplifier_w = total_available_optical_w / (
            repeater_count[i] * amplifiers_per_repeater
        )

        valid = (
            (launch_powers_w <= pfe_cap_per_amplifier_w)
            & (launch_powers_w <= nonlinear_cap_w)
            & (launch_powers_w <= hardware_cap_w)
        )
        if not np.any(valid):
            continue

        feasible_indices = np.flatnonzero(valid)
        local_opt = int(np.nanargmax(snr_curve[valid]))
        global_opt = int(feasible_indices[local_opt])
        feasible_span[i] = True
        max_snr_linear[i] = snr_curve[global_opt]
        max_power_w[i] = launch_powers_w[global_opt]
        eta_at_optimum[i] = efficiency[i, global_opt]

    # Shannon-like capacity with optional implementation gap and net overhead.
    effective_snr = max_snr_linear / implementation_gap_linear
    capacity_bps = (
        M
        * n_cores
        * n_channels
        * Rs
        * np.log2(1.0 + effective_snr)
        * (1.0 - overhead_fraction)
    )

    allowed = feasible_span & (span_candidates >= L_span_min)
    if not np.any(allowed):
        raise ValueError("No feasible span candidate satisfies L_span_min.")
    allowed_indices = np.flatnonzero(allowed)
    best_local = int(np.nanargmax(capacity_bps[allowed]))
    best_idx = int(allowed_indices[best_local])

    result = CapacityOptimizationResult(
        eta_at_optimum=eta_at_optimum,
        optimal_span_length_km=float(span_candidates[best_idx]),
        maximum_capacity_bps=float(capacity_bps[best_idx]),
        gsnr_at_capacity_optimum_linear=float(max_snr_linear[best_idx]),
        gsnr_by_span_linear=max_snr_linear,
        launch_power_by_span_w=max_power_w,
        repeater_count=repeater_count,
        span_count=span_count,
        unconstrained_nonlinear_threshold_w=unconstrained_threshold_w,
        feasible=feasible_span,
    )

    if return_details:
        return result

    # Backward-compatible tuple, but SNR_max now corresponds to C_max.
    return (
        result.eta_at_optimum,
        result.optimal_span_length_km,
        result.maximum_capacity_bps,
        result.gsnr_at_capacity_optimum_linear,
        result.gsnr_by_span_linear,
        result.launch_power_by_span_w,
    )


# Preferred PEP-8 alias.
optimize_capacity_extended = optimized_capacity_extended
