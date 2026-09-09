"""Reusable GN-model engine and validation CLI.

This file is a reusable version of the attached G.654.E notebook/report code.
It keeps the same closed-form GN approximation and the same default conditions,
but removes the notebook-only fixed input block and the image-specific fitting.

Default conventions:
  - launch power is total dual-polarization channel power
  - GN coefficient is 8/27
  - NLI accumulation is incoherent and proportional to span count
  - ASE bandwidth is the symbol rate, matching the attached code
  - one loss-compensating EDFA is used per span

The 8/27 coefficient is not silently changed. If another implementation uses
4/27, 16/27, per-polarization power, or a different ASE bandwidth, pass those
choices explicitly and record them in the returned metadata.

Typical import:

    from main import (
        FiberParameters, SystemParameters, GNOptions,
        simulate_snr_curve, validate_gn_model,
    )

    fiber = FiberParameters(
        name="G.654.E",
        attenuation_db_per_km=0.166,
        effective_area_um2=125.0,
        dispersion_ps_nm_km=21.0,
        n2_m2_w=2.2e-20,
    )
    system = SystemParameters(
        channels=90,
        symbol_rate_gbd=95.0,
        spacing_ghz=95.0,
        span_length_km=80.0,
        noise_figure_db=5.0,
        transceiver_snr_db=18.0,
    )
    options = GNOptions(nli_coefficient=8/27, finite_effective_length=True)

    result = validate_gn_model(
        launch_dbm=[-10, -5, 0, 4, 8, 10],
        spans=30,
        fiber=fiber,
        system=system,
        options=options,
        reference_snr_db=external_snr_db,
    )
    print(result["validation"]["metrics"])

CLI:
    python main.py --model finite --spans 1,10,30,50 --save-dir gn_results
    python main.py --reference-csv other_model.csv --spans 30

The reference CSV must contain launch_dbm and snr_db columns. If it contains
spans, the selected span is used.
"""

from __future__ import annotations

import argparse
import csv
import json
import platform
from dataclasses import asdict, dataclass, replace
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, Mapping, Optional, Sequence, Tuple

import numpy as np

try:
    import pandas as pd
except ImportError:  # pandas is optional for library use
    pd = None


# ---------------------------------------------------------------------------
# Input models
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class FiberParameters:
    """Fiber parameters.

    alpha_power_per_km, beta2_s2_per_km, and gamma_per_w_km are optional direct
    overrides. When omitted, they are derived from the usual fiber inputs.
    """

    name: str = "G.654.E"
    wavelength_nm: float = 1550.0
    attenuation_db_per_km: float = 0.166
    effective_area_um2: float = 125.0
    dispersion_ps_nm_km: float = 21.0
    n2_m2_w: float = 2.2e-20
    alpha_power_per_km: Optional[float] = None
    beta2_s2_per_km: Optional[float] = None
    gamma_per_w_km: Optional[float] = None


@dataclass(frozen=True)
class SystemParameters:
    """System and transceiver parameters.

    The attached code uses symbol rate as the ASE noise bandwidth. That
    behavior is retained by leaving ase_bandwidth_hz as None.
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
class GNOptions:
    """GN-model conventions.

    nli_coefficient=8/27 is the default used in the attached code for total
    dual-polarization channel power. The code does not infer a different
    coefficient from polarizations, because that would hide a convention
    mismatch during validation.
    """

    nli_coefficient: float = 8.0 / 27.0
    finite_effective_length: bool = True
    power_definition: str = "total_dp"
    nli_channel_count: Optional[float] = None
    model_name: str = "gn_finite"


C0 = 299_792_458.0
H_PLANCK = 6.626_070_15e-34
DB_PER_NEPER_POWER = 10.0 / np.log(10.0)


# ---------------------------------------------------------------------------
# Unit conversion and physical model
# ---------------------------------------------------------------------------


def _as_float_array(values: Any) -> np.ndarray:
    array = np.atleast_1d(np.asarray(values, dtype=float))
    if not np.all(np.isfinite(array)):
        raise ValueError("입력 배열에 유한하지 않은 값이 있습니다.")
    return array


def _validate_fiber(fiber: FiberParameters) -> None:
    for name, value in {
        "wavelength_nm": fiber.wavelength_nm,
        "effective_area_um2": fiber.effective_area_um2,
        "n2_m2_w": fiber.n2_m2_w,
    }.items():
        if value <= 0:
            raise ValueError(f"{name}은(는) 0보다 커야 합니다.")
    if fiber.alpha_power_per_km is not None and fiber.alpha_power_per_km <= 0:
        raise ValueError("alpha_power_per_km은 0보다 커야 합니다.")
    if fiber.beta2_s2_per_km is not None and fiber.beta2_s2_per_km == 0:
        raise ValueError("beta2_s2_per_km은 0이 될 수 없습니다.")
    if fiber.gamma_per_w_km is not None and fiber.gamma_per_w_km <= 0:
        raise ValueError("gamma_per_w_km은 0보다 커야 합니다.")


def _validate_system(system: SystemParameters) -> None:
    if isinstance(system.channels, bool) or system.channels < 1:
        raise ValueError("channels는 1 이상의 정수여야 합니다.")
    for name, value in {
        "symbol_rate_gbd": system.symbol_rate_gbd,
        "spacing_ghz": system.spacing_ghz,
        "span_length_km": system.span_length_km,
        "noise_figure_db": system.noise_figure_db,
        "transceiver_snr_db": system.transceiver_snr_db,
    }.items():
        if not np.isfinite(value) or value <= 0:
            raise ValueError(f"{name}은(는) 0보다 큰 유한한 값이어야 합니다.")
    if system.polarizations not in (1, 2):
        raise ValueError("polarizations는 1 또는 2로 지정하세요.")
    if system.ase_bandwidth_hz is not None and system.ase_bandwidth_hz <= 0:
        raise ValueError("ase_bandwidth_hz는 0보다 커야 합니다.")


def _validate_options(options: GNOptions) -> None:
    if not np.isfinite(options.nli_coefficient) or options.nli_coefficient <= 0:
        raise ValueError("nli_coefficient는 0보다 큰 유한한 값이어야 합니다.")
    if options.power_definition not in {"total_dp", "per_polarization", "single_pol"}:
        raise ValueError(
            "power_definition은 total_dp, per_polarization, single_pol 중 하나여야 합니다."
        )
    if options.nli_channel_count is not None and options.nli_channel_count <= 0:
        raise ValueError("nli_channel_count는 0보다 커야 합니다.")


def derive_parameters(
    fiber: FiberParameters = FiberParameters(),
    system: SystemParameters = SystemParameters(),
) -> Dict[str, float]:
    """Convert fiber/system values into GN parameters with explicit units.

    Returned units:
      alpha_power_per_km: km^-1
      beta2_s2_per_km: s^2/km
      gamma_per_w_km: W^-1/km
    """

    _validate_fiber(fiber)
    _validate_system(system)

    wavelength_m = fiber.wavelength_nm * 1e-9
    rate_hz = system.symbol_rate_gbd * 1e9
    spacing_hz = system.spacing_ghz * 1e9

    if fiber.alpha_power_per_km is None:
        alpha_power_per_km = fiber.attenuation_db_per_km / DB_PER_NEPER_POWER
        span_gain_linear = 10.0 ** (
            fiber.attenuation_db_per_km * system.span_length_km / 10.0
        )
        attenuation_db_per_km = fiber.attenuation_db_per_km
    else:
        alpha_power_per_km = float(fiber.alpha_power_per_km)
        span_gain_linear = np.exp(alpha_power_per_km * system.span_length_km)
        attenuation_db_per_km = DB_PER_NEPER_POWER * alpha_power_per_km

    if fiber.beta2_s2_per_km is None:
        # 1 ps/(nm km) = 1e-6 s/m^2.
        dispersion_si = fiber.dispersion_ps_nm_km * 1e-6
        beta2_s2_per_km = -(
            wavelength_m**2 / (2.0 * np.pi * C0)
        ) * dispersion_si * 1000.0
    else:
        beta2_s2_per_km = float(fiber.beta2_s2_per_km)

    if fiber.gamma_per_w_km is None:
        gamma_per_w_km = (
            2.0
            * np.pi
            * fiber.n2_m2_w
            / (wavelength_m * fiber.effective_area_um2 * 1e-12)
            * 1000.0
        )
    else:
        gamma_per_w_km = float(fiber.gamma_per_w_km)

    effective_length_km = (
        -np.expm1(-alpha_power_per_km * system.span_length_km)
        / alpha_power_per_km
    )
    asymptotic_length_km = 1.0 / alpha_power_per_km
    noise_figure_linear = 10.0 ** (system.noise_figure_db / 10.0)
    ase_bandwidth_hz = (
        rate_hz if system.ase_bandwidth_hz is None else system.ase_bandwidth_hz
    )
    ase_per_span_w = (
        H_PLANCK
        * (C0 / wavelength_m)
        * noise_figure_linear
        * ase_bandwidth_hz
        * (span_gain_linear - 1.0)
    )
    occupied_bandwidth_thz = system.channels * system.spacing_ghz / 1000.0
    channels_that_fit = (
        None
        if system.stated_gain_bandwidth_thz is None
        else int(system.stated_gain_bandwidth_thz * 1000.0 // system.spacing_ghz)
    )

    return {
        "wavelength_m": float(wavelength_m),
        "frequency_hz": float(C0 / wavelength_m),
        "rate_hz": float(rate_hz),
        "spacing_hz": float(spacing_hz),
        "alpha_power_per_km": float(alpha_power_per_km),
        "alpha_field_per_km": float(alpha_power_per_km / 2.0),
        "beta2_s2_per_km": float(beta2_s2_per_km),
        "beta2_ps2_per_km": float(beta2_s2_per_km * 1e24),
        "gamma_per_w_km": float(gamma_per_w_km),
        "effective_length_km": float(effective_length_km),
        "asymptotic_length_km": float(asymptotic_length_km),
        "attenuation_db_per_km_used": float(attenuation_db_per_km),
        "span_gain_db": float(10.0 * np.log10(span_gain_linear)),
        "span_gain_linear": float(span_gain_linear),
        "ase_bandwidth_hz": float(ase_bandwidth_hz),
        "ase_per_span_w": float(ase_per_span_w),
        "wdm_occupied_bandwidth_thz": float(occupied_bandwidth_thz),
        "channels_that_fit_stated_bandwidth": channels_that_fit,
    }


def gn_eta_per_span(
    parameters: Mapping[str, float],
    system: SystemParameters,
    options: GNOptions = GNOptions(),
    *,
    finite: Optional[bool] = None,
    n_channels: Optional[float] = None,
) -> float:
    """Return the per-span NLI coefficient eta in W^-2.

    This is the same central-channel closed-form approximation used in the
    attached code. n_channels lets a caller validate another channel plan.
    """

    _validate_system(system)
    _validate_options(options)
    finite = options.finite_effective_length if finite is None else finite
    beta_abs = abs(float(parameters["beta2_s2_per_km"]))
    if beta_abs == 0:
        raise ValueError("GN approximation requires nonzero absolute beta2.")

    if n_channels is None:
        channel_count = (
            system.channels
            if options.nli_channel_count is None
            else options.nli_channel_count
        )
    else:
        channel_count = n_channels
    if channel_count <= 0:
        raise ValueError("n_channels는 0보다 커야 합니다.")

    beta_length = float(parameters["asymptotic_length_km"])
    effective_length = (
        float(parameters["effective_length_km"]) if finite else beta_length
    )
    rate_hz = float(parameters["rate_hz"])
    spacing_hz = float(parameters["spacing_hz"])
    argument = (
        (np.pi**2 / 2.0)
        * beta_abs
        * beta_length
        * rate_hz**2
        * channel_count ** (2.0 * rate_hz / spacing_hz)
    )
    eta = (
        options.nli_coefficient
        * float(parameters["gamma_per_w_km"]) ** 2
        * effective_length**2
        / (np.pi * beta_abs * beta_length * rate_hz**2)
        * np.arcsinh(argument)
    )
    return float(eta)


def _validate_spans(spans: int) -> int:
    if isinstance(spans, bool) or int(spans) != spans or spans < 1:
        raise ValueError("spans는 1 이상의 정수여야 합니다.")
    return int(spans)


def calculate_snr(
    launch_dbm: Sequence[float] | np.ndarray | float,
    spans: int,
    eta_per_span: float,
    parameters: Mapping[str, float],
    system: SystemParameters,
) -> Dict[str, np.ndarray]:
    """Calculate signal, ASE, NLI, transceiver noise, and end-to-end SNR."""

    _validate_system(system)
    spans = _validate_spans(spans)
    launch = _as_float_array(launch_dbm)
    if not np.isfinite(eta_per_span) or eta_per_span < 0:
        raise ValueError("eta_per_span은 0 이상인 유한한 값이어야 합니다.")

    signal_w = 1e-3 * 10.0 ** (launch / 10.0)
    ase_w = np.full_like(signal_w, spans * float(parameters["ase_per_span_w"]))
    nli_w = spans * eta_per_span * signal_w**3
    trx_equivalent_noise_w = signal_w / 10.0 ** (system.transceiver_snr_db / 10.0)
    total_noise_w = ase_w + nli_w + trx_equivalent_noise_w
    snr_linear = signal_w / total_noise_w

    return {
        "launch_dbm": launch,
        "spans": np.full(launch.shape, spans, dtype=int),
        "distance_km": np.full(launch.shape, spans * system.span_length_km),
        "signal_w": signal_w,
        "ase_w": ase_w,
        "nli_w": nli_w,
        "trx_equivalent_noise_w": trx_equivalent_noise_w,
        "total_noise_w": total_noise_w,
        "snr_linear": snr_linear,
        "snr_db": 10.0 * np.log10(snr_linear),
    }


def records_to_dataframe(columns: Mapping[str, Sequence[Any]]):
    """Return a DataFrame when pandas is available, otherwise raw arrays."""

    if pd is None:
        return {key: np.asarray(value) for key, value in columns.items()}
    return pd.DataFrame(columns)


def _concat_result_rows(rows: Iterable[Mapping[str, Any]]):
    rows = list(rows)
    if not rows:
        return records_to_dataframe({})
    if pd is None:
        return rows
    return pd.DataFrame(rows)


def _result_to_rows(result: Mapping[str, np.ndarray]) -> list[dict[str, Any]]:
    keys = list(result)
    size = len(result[keys[0]]) if keys else 0
    return [
        {
            key: (
                value[index].item()
                if isinstance(value[index], np.generic)
                else value[index]
            )
            for key, value in result.items()
        }
        for index in range(size)
    ]


def model_eta(
    fiber: FiberParameters = FiberParameters(),
    system: SystemParameters = SystemParameters(),
    options: GNOptions = GNOptions(),
    *,
    model: Optional[str] = None,
) -> Tuple[float, Dict[str, float]]:
    """Return eta and derived parameters for legacy or finite-length GN."""

    if model is not None:
        if model not in {"legacy", "finite", "gn_finite"}:
            raise ValueError("model은 legacy, finite, gn_finite 중 하나여야 합니다.")
        options = replace(
            options,
            finite_effective_length=(model != "legacy"),
            model_name="legacy" if model == "legacy" else "gn_finite",
        )
    parameters = derive_parameters(fiber, system)
    eta = gn_eta_per_span(parameters, system, options)
    return eta, parameters


def simulate_snr_curve(
    launch_dbm: Sequence[float] | np.ndarray | float,
    spans: int,
    fiber: FiberParameters = FiberParameters(),
    system: SystemParameters = SystemParameters(),
    options: GNOptions = GNOptions(),
    *,
    model: Optional[str] = None,
):
    """Run the GN model and return one row per launch-power point."""

    eta, _ = model_eta(fiber, system, options, model=model)
    result = calculate_snr(
        launch_dbm,
        spans,
        eta,
        derive_parameters(fiber, system),
        system,
    )
    rows = _result_to_rows(result)
    for row in rows:
        row["model"] = model if model is not None else options.model_name
        row["eta_per_span_w_inv2"] = eta
    return _concat_result_rows(rows)


def analytical_optimum_launch_dbm(
    eta_per_span: float,
    parameters: Mapping[str, float],
) -> float:
    """Return the stationary launch power for ASE + NLI + TRX noise."""

    if eta_per_span <= 0:
        raise ValueError("eta_per_span은 0보다 커야 합니다.")
    power_w = (float(parameters["ase_per_span_w"]) / (2.0 * eta_per_span)) ** (
        1.0 / 3.0
    )
    return float(10.0 * np.log10(power_w / 1e-3))


def make_launch_grid(
    launch_min_dbm: float = -10.0,
    launch_max_dbm: float = 10.0,
    launch_step_db: float = 0.05,
) -> np.ndarray:
    """Create an inclusive launch-power grid."""

    if launch_max_dbm <= launch_min_dbm or launch_step_db <= 0:
        raise ValueError(
            "launch_min_dbm < launch_max_dbm, launch_step_db > 0 이어야 합니다."
        )
    number = int(np.floor((launch_max_dbm - launch_min_dbm) / launch_step_db))
    grid = launch_min_dbm + np.arange(number + 1) * launch_step_db
    if grid[-1] < launch_max_dbm - 1e-10:
        grid = np.append(grid, launch_max_dbm)
    return grid


# ---------------------------------------------------------------------------
# Independent validation against another model, measurement, or CSV
# ---------------------------------------------------------------------------


def error_metrics(
    reference_snr_db: Sequence[float] | np.ndarray,
    predicted_snr_db: Sequence[float] | np.ndarray,
) -> Dict[str, float]:
    """Calculate dB errors and linear-SNR MAPE."""

    reference = _as_float_array(reference_snr_db)
    predicted = _as_float_array(predicted_snr_db)
    if reference.shape != predicted.shape:
        raise ValueError("reference와 predicted의 배열 길이가 다릅니다.")
    delta_db = predicted - reference
    return {
        "n_points": int(delta_db.size),
        "mae_db": float(np.mean(np.abs(delta_db))),
        "rmse_db": float(np.sqrt(np.mean(delta_db**2))),
        "max_abs_error_db": float(np.max(np.abs(delta_db))),
        "mean_bias_db": float(np.mean(delta_db)),
        "linear_snr_mape_pct": float(
            100.0 * np.mean(np.abs(10.0 ** (delta_db / 10.0) - 1.0))
        ),
    }


def compare_curves(
    reference_launch_dbm: Sequence[float] | np.ndarray,
    reference_snr_db: Sequence[float] | np.ndarray,
    predicted_launch_dbm: Sequence[float] | np.ndarray,
    predicted_snr_db: Sequence[float] | np.ndarray,
    *,
    interpolate: bool = True,
):
    """Compare two SNR curves on a common launch-power grid."""

    ref_x = _as_float_array(reference_launch_dbm)
    ref_y = _as_float_array(reference_snr_db)
    pred_x = _as_float_array(predicted_launch_dbm)
    pred_y = _as_float_array(predicted_snr_db)
    if ref_x.size != ref_y.size or pred_x.size != pred_y.size:
        raise ValueError("각 curve의 x/y 길이가 일치해야 합니다.")

    order_ref = np.argsort(ref_x)
    order_pred = np.argsort(pred_x)
    ref_x, ref_y = ref_x[order_ref], ref_y[order_ref]
    pred_x, pred_y = pred_x[order_pred], pred_y[order_pred]
    if interpolate:
        mask = (ref_x >= pred_x[0]) & (ref_x <= pred_x[-1])
        x = ref_x[mask]
        reference = ref_y[mask]
        predicted = np.interp(x, pred_x, pred_y)
    else:
        if ref_x.shape != pred_x.shape or not np.allclose(ref_x, pred_x):
            raise ValueError(
                "interpolate=False이면 두 launch-power grid가 같아야 합니다."
            )
        x, reference, predicted = ref_x, ref_y, pred_y

    if x.size == 0:
        raise ValueError("두 curve의 launch-power 범위가 겹치지 않습니다.")

    comparison = records_to_dataframe(
        {
            "launch_dbm": x,
            "reference_snr_db": reference,
            "predicted_snr_db": predicted,
            "error_db": predicted - reference,
            "absolute_error_db": np.abs(predicted - reference),
        }
    )
    return {
        "comparison": comparison,
        "metrics": error_metrics(reference, predicted),
    }


def validate_gn_model(
    launch_dbm: Sequence[float] | np.ndarray,
    spans: int,
    *,
    fiber: FiberParameters = FiberParameters(),
    system: SystemParameters = SystemParameters(),
    options: GNOptions = GNOptions(),
    model: Optional[str] = None,
    reference_snr_db: Optional[Sequence[float] | np.ndarray] = None,
    reference_launch_dbm: Optional[Sequence[float] | np.ndarray] = None,
) -> Dict[str, Any]:
    """Run this GN model and optionally compare it with another result.

    reference_snr_db can be the output of another GN/GGN/SSF code or a
    measured curve. If reference_launch_dbm is omitted, launch_dbm is used.
    """

    launch = _as_float_array(launch_dbm)
    curve = simulate_snr_curve(
        launch,
        spans,
        fiber=fiber,
        system=system,
        options=options,
        model=model,
    )
    eta, parameters = model_eta(fiber, system, options, model=model)
    result: Dict[str, Any] = {
        "curve": curve,
        "eta_per_span_w_inv2": eta,
        "parameters": parameters,
        "fiber": asdict(fiber),
        "system": asdict(system),
        "options": asdict(options),
    }

    if reference_snr_db is not None:
        ref_x = launch if reference_launch_dbm is None else reference_launch_dbm
        predicted_x = (
            curve["launch_dbm"].to_numpy()
            if pd is not None
            else np.asarray(curve["launch_dbm"])
        )
        predicted_y = (
            curve["snr_db"].to_numpy()
            if pd is not None
            else np.asarray(curve["snr_db"])
        )
        result["validation"] = compare_curves(
            ref_x,
            reference_snr_db,
            predicted_x,
            predicted_y,
        )
    return result


def validate_against_callable(
    launch_dbm: Sequence[float] | np.ndarray,
    spans: int,
    reference_function: Callable[..., Sequence[float]],
    *,
    fiber: FiberParameters = FiberParameters(),
    system: SystemParameters = SystemParameters(),
    options: GNOptions = GNOptions(),
    model: Optional[str] = None,
) -> Dict[str, Any]:
    """Validate against a second Python implementation.

    The reference function is called as:
        reference_function(launch_dbm, spans, fiber, system)

    Wrap another implementation in an adapter if it needs a different
    argument signature.
    """

    launch = _as_float_array(launch_dbm)
    reference = _as_float_array(reference_function(launch, spans, fiber, system))
    return validate_gn_model(
        launch,
        spans,
        fiber=fiber,
        system=system,
        options=options,
        model=model,
        reference_snr_db=reference,
    )


def fit_eta_scale_to_reference(
    reference_launch_dbm: Sequence[float] | np.ndarray,
    reference_snr_db: Sequence[float] | np.ndarray,
    spans: int,
    *,
    fiber: FiberParameters = FiberParameters(),
    system: SystemParameters = SystemParameters(),
    options: GNOptions = GNOptions(),
    log10_scale_bounds: Tuple[float, float] = (-4.0, 2.0),
) -> Dict[str, Any]:
    """Optional empirical eta fit, separate from independent validation.

    A fitted scale measures agreement with the supplied curve. It is not
    independent proof of the GN implementation.
    """

    x = _as_float_array(reference_launch_dbm)
    y = _as_float_array(reference_snr_db)
    eta, parameters = model_eta(fiber, system, options)
    lo, hi = log10_scale_bounds
    if hi <= lo:
        raise ValueError("log10_scale_bounds가 올바르지 않습니다.")

    def objective(log_scale: float) -> float:
        predicted = calculate_snr(
            x,
            spans,
            eta * 10.0**log_scale,
            parameters,
            system,
        )["snr_db"]
        return error_metrics(y, predicted)["rmse_db"]

    # Coarse scan and golden-section refinement keep scipy optional.
    scan = np.linspace(lo, hi, 401)
    values = np.asarray([objective(value) for value in scan])
    best_index = int(np.argmin(values))
    left = scan[max(0, best_index - 1)]
    right = scan[min(len(scan) - 1, best_index + 1)]
    phi = (1.0 + np.sqrt(5.0)) / 2.0
    for _ in range(80):
        c = right - (right - left) / phi
        d = left + (right - left) / phi
        if objective(c) < objective(d):
            right = d
        else:
            left = c
    best_log_scale = (left + right) / 2.0
    fitted_eta = eta * 10.0**best_log_scale
    fitted = calculate_snr(x, spans, fitted_eta, parameters, system)["snr_db"]
    return {
        "scale": float(10.0**best_log_scale),
        "base_eta_per_span_w_inv2": float(eta),
        "fitted_eta_per_span_w_inv2": float(fitted_eta),
        "comparison": compare_curves(x, y, x, fitted, interpolate=False),
        "warning": "동일 reference에 대한 경험적 보정이며 독립 검증값이 아닙니다.",
    }


# ---------------------------------------------------------------------------
# CSV, JSON, and command-line helpers
# ---------------------------------------------------------------------------


def read_reference_csv(
    path: str | Path,
    *,
    spans: Optional[int] = None,
    launch_column: str = "launch_dbm",
    snr_column: str = "snr_db",
    spans_column: str = "spans",
) -> Tuple[np.ndarray, np.ndarray]:
    """Read another model's or measurement's curve from CSV."""

    path = Path(path)
    if pd is not None:
        frame = pd.read_csv(path)
        if spans is not None and spans_column in frame.columns:
            frame = frame.loc[frame[spans_column] == spans]
        if launch_column not in frame.columns or snr_column not in frame.columns:
            raise ValueError(f"CSV에는 {launch_column}, {snr_column} 열이 필요합니다.")
        return (
            frame[launch_column].to_numpy(dtype=float),
            frame[snr_column].to_numpy(dtype=float),
        )

    with path.open(newline="", encoding="utf-8-sig") as handle:
        rows = list(csv.DictReader(handle))
    if spans is not None and rows and spans_column in rows[0]:
        rows = [row for row in rows if int(float(row[spans_column])) == spans]
    return (
        np.asarray([float(row[launch_column]) for row in rows]),
        np.asarray([float(row[snr_column]) for row in rows]),
    )


def save_result(
    result: Mapping[str, Any],
    directory: str | Path,
    *,
    prefix: str = "gn_model",
) -> Path:
    """Save curve, validation table, parameters, and metadata."""

    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    curve = result.get("curve")
    curve_path = directory / f"{prefix}_curve.csv"
    if pd is not None and hasattr(curve, "to_csv"):
        curve.to_csv(curve_path, index=False, encoding="utf-8-sig")
    else:
        _write_mapping_csv(curve, curve_path)

    if "validation" in result:
        comparison = result["validation"]["comparison"]
        comparison_path = directory / f"{prefix}_validation.csv"
        if pd is not None and hasattr(comparison, "to_csv"):
            comparison.to_csv(comparison_path, index=False, encoding="utf-8-sig")
        else:
            _write_mapping_csv(comparison, comparison_path)
        with (directory / f"{prefix}_metrics.json").open("w", encoding="utf-8") as handle:
            json.dump(result["validation"]["metrics"], handle, indent=2, ensure_ascii=False)

    metadata = {
        "fiber": result.get("fiber"),
        "system": result.get("system"),
        "options": result.get("options"),
        "parameters": result.get("parameters"),
        "eta_per_span_w_inv2": result.get("eta_per_span_w_inv2"),
        "python": platform.python_version(),
        "numpy": np.__version__,
    }
    with (directory / f"{prefix}_metadata.json").open("w", encoding="utf-8") as handle:
        json.dump(metadata, handle, indent=2, ensure_ascii=False, default=_json_default)
    return directory


def _write_mapping_csv(data: Any, path: Path) -> None:
    if isinstance(data, Mapping):
        keys = list(data)
        rows = zip(*[np.asarray(data[key]).tolist() for key in keys])
        with path.open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.writer(handle)
            writer.writerow(keys)
            writer.writerows(rows)
        return
    if isinstance(data, list):
        keys = list(data[0]) if data else []
        with path.open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.DictWriter(handle, fieldnames=keys)
            writer.writeheader()
            writer.writerows(data)
        return
    raise TypeError("CSV로 저장할 수 없는 결과 형식입니다.")


def _json_default(value: Any):
    if isinstance(value, np.generic):
        return value.item()
    if isinstance(value, np.ndarray):
        return value.tolist()
    raise TypeError(f"JSON으로 변환할 수 없는 형식: {type(value)}")


def _plot_result(result: Mapping[str, Any], path: Optional[Path] = None) -> None:
    try:
        import matplotlib.pyplot as plt
    except ImportError as exc:
        raise RuntimeError("그래프에는 matplotlib가 필요합니다.") from exc

    curve = result["curve"]
    x = curve["launch_dbm"].to_numpy() if pd is not None else curve["launch_dbm"]
    y = curve["snr_db"].to_numpy() if pd is not None else curve["snr_db"]
    fig, ax = plt.subplots(figsize=(9, 5.5), layout="constrained")
    ax.plot(x, y, lw=2.2, label="GN model")
    if "validation" in result:
        comparison = result["validation"]["comparison"]
        cx = (
            comparison["launch_dbm"].to_numpy()
            if pd is not None
            else comparison["launch_dbm"]
        )
        ry = (
            comparison["reference_snr_db"].to_numpy()
            if pd is not None
            else comparison["reference_snr_db"]
        )
        ax.plot(cx, ry, "o", ms=3, alpha=0.7, label="reference")
    ax.set_xlabel("Launch power per channel, total DP (dBm)")
    ax.set_ylabel("End-to-end SNR (dB)")
    ax.grid(alpha=0.3)
    ax.legend()
    if path is None:
        plt.show()
    else:
        fig.savefig(path, dpi=180)
        plt.close(fig)


def run_cli(argv: Optional[Sequence[str]] = None) -> Dict[str, Any]:
    parser = argparse.ArgumentParser(
        description="Reusable GN-model calculation and validation"
    )
    parser.add_argument("--model", choices=["legacy", "finite"], default="finite")
    parser.add_argument("--spans", default="1,10,30,50")
    parser.add_argument("--launch-min", type=float, default=-10.0)
    parser.add_argument("--launch-max", type=float, default=10.0)
    parser.add_argument("--launch-step", type=float, default=0.05)
    parser.add_argument("--coefficient", type=float, default=8.0 / 27.0)
    parser.add_argument("--reference-csv", type=str, default=None)
    parser.add_argument("--save-dir", type=str, default="gn_results")
    parser.add_argument("--plot", action="store_true")
    args = parser.parse_args(argv)

    spans_list = [
        int(value.strip()) for value in args.spans.split(",") if value.strip()
    ]
    launch = make_launch_grid(args.launch_min, args.launch_max, args.launch_step)
    fiber = FiberParameters()
    system = SystemParameters()
    options = GNOptions(
        nli_coefficient=args.coefficient,
        finite_effective_length=args.model == "finite",
        model_name="gn_finite" if args.model == "finite" else "legacy",
    )

    results = {}
    for spans in spans_list:
        reference_x = reference_y = None
        if args.reference_csv is not None:
            reference_x, reference_y = read_reference_csv(
                args.reference_csv, spans=spans
            )
            run_launch = reference_x
        else:
            run_launch = launch
        result = validate_gn_model(
            run_launch,
            spans,
            fiber=fiber,
            system=system,
            options=options,
            reference_snr_db=reference_y,
            reference_launch_dbm=reference_x,
        )
        results[spans] = result
        save_result(
            result,
            Path(args.save_dir) / f"spans_{spans}",
            prefix="gn",
        )
        if args.plot:
            _plot_result(
                result,
                Path(args.save_dir) / f"spans_{spans}" / "gn_curve.png",
            )
        if "validation" in result:
            print(f"spans={spans}: {result['validation']['metrics']}")
        else:
            print(f"spans={spans}: eta={result['eta_per_span_w_inv2']:.6e}")
    return results


def main(argv: Optional[Sequence[str]] = None) -> Dict[str, Any]:
    """CLI-compatible entry point."""

    return run_cli(argv)


if __name__ == "__main__":
    main()
