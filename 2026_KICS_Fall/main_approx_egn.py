"""Approximate EGN extension for the reusable main.py GN model.

This module deliberately keeps main.py as the GN baseline and adds the
closed-form approximate EGN (App. EGN) correction from the attached paper:

    P. Poggiolini et al., "A Simple and Accurate Closed-Form EGN Model
    Formula", arXiv:1503.04132v1, 2015.

The implemented equations are Eq. (1) and the identical/equally-spaced,
centre-channel special case in Eq. (3):

    G_NLI^EGN ~= G_NLI^GN - G_corr

This is NOT a full numerical EGN integral solver. It is a lightweight,
closed-form correction to the existing GN result. It makes no fit to a
reference curve: the fibre, WDM, modulation, span, and launch-power inputs
are the only inputs that set the result.

Supported modulation factors from the attached paper:
  - PM-QPSK: phi = 1
  - PM-16QAM: phi = 17/25
  - PM-64QAM: phi = 13/21

The correction assumes lumped amplification, non-zero loss, identical
equally-spaced channels, ideally rectangular channel spectra, and a
centre-channel under test. For an even number of channels, this script uses
the analytical continuation of the harmonic number for a pseudo-centre
channel and records that approximation in the result metadata.

Typical use:

    python main_approx_egn.py --spans 30 --modulation PM-16QAM --plot
    python main_approx_egn.py --channels 9 --symbol-rate-gbd 32
        --spacing-ghz 33.6 --span-length-km 100 --modulation PM-QPSK

Programmatic use:

    from main import FiberParameters, SystemParameters, GNOptions
    from main_approx_egn import validate_approximate_egn_model

    result = validate_approximate_egn_model(
        launch_dbm=[-6, -3, 0, 3],
        spans=30,
        fiber=FiberParameters(),
        system=SystemParameters(),
        gn_options=GNOptions(nli_coefficient=8/27),
        egn_options=ApproxEGNOptions(modulation_format="PM-64QAM"),
    )
"""

from __future__ import annotations

import argparse
import csv
import json
import platform
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Dict, Mapping, Optional, Sequence

import numpy as np

from main import (
    FiberParameters,
    GNOptions,
    SystemParameters,
    calculate_snr,
    compare_curves,
    derive_parameters,
    gn_eta_per_span,
    make_launch_grid,
    read_reference_csv,
    records_to_dataframe,
)


PAPER_REFERENCE = (
    "P. Poggiolini et al., A Simple and Accurate Closed-Form EGN Model "
    "Formula, arXiv:1503.04132v1 (2015)."
)
PAPER_EQUATIONS = "Eq. (1) and Eq. (3), closed-form approximate EGN"
EULER_MASCHERONI = 0.5772156649015329

# Eq. (2)/(3) modulation-dependent factors from the attached paper.
MODULATION_PHI: Dict[str, float] = {
    "PM-QPSK": 1.0,
    "PM-16QAM": 17.0 / 25.0,
    "PM-64QAM": 13.0 / 21.0,
}


@dataclass(frozen=True)
class ApproxEGNOptions:
    """Options specific to the closed-form approximate EGN correction."""

    modulation_format: str = "PM-64QAM"
    strict_applicability: bool = True
    allow_even_channel_center_approximation: bool = True
    model_name: str = "approx_egn_closed_form"


def _canonical_modulation(modulation_format: str) -> str:
    """Return a supported canonical PM modulation name."""

    key = "".join(char for char in str(modulation_format).upper() if char.isalnum())
    aliases = {
        "QPSK": "PM-QPSK",
        "PMQPSK": "PM-QPSK",
        "16QAM": "PM-16QAM",
        "PM16QAM": "PM-16QAM",
        "64QAM": "PM-64QAM",
        "PM64QAM": "PM-64QAM",
    }
    if key not in aliases:
        choices = ", ".join(MODULATION_PHI)
        raise ValueError(
            f"지원하지 않는 EGN 변조 형식: {modulation_format!r}. "
            f"지원 형식: {choices}"
        )
    return aliases[key]


def modulation_phi(modulation_format: str) -> float:
    """Return the paper's modulation-dependent Phi coefficient."""

    return MODULATION_PHI[_canonical_modulation(modulation_format)]


def _digamma_positive(value: float) -> float:
    """Accurate digamma approximation for a positive real value.

    It avoids making SciPy a dependency and is used only for even-channel
    pseudo-centre cases, where a half-integer harmonic number is needed.
    """

    if value <= 0.0:
        raise ValueError("digamma 입력은 0보다 커야 합니다.")
    result = 0.0
    x = float(value)
    while x < 8.0:
        result -= 1.0 / x
        x += 1.0
    inverse = 1.0 / x
    inverse_squared = inverse * inverse
    return float(
        result
        + np.log(x)
        - 0.5 * inverse
        - inverse_squared
        * (
            1.0 / 12.0
            - inverse_squared
            * (
                1.0 / 120.0
                - inverse_squared * (1.0 / 252.0 - inverse_squared / 240.0)
            )
        )
    )


def harmonic_number(order: float) -> float:
    """Return H_order, including a half-integer continuation when needed."""

    x = float(order)
    if x < 0.0:
        raise ValueError("Harmonic number order는 0 이상이어야 합니다.")
    if np.isclose(x, 0.0, atol=1e-14):
        return 0.0
    nearest_integer = int(round(x))
    if np.isclose(x, nearest_integer, rtol=0.0, atol=1e-12):
        return float(np.sum(1.0 / np.arange(1, nearest_integer + 1, dtype=float)))
    return float(_digamma_positive(x + 1.0) + EULER_MASCHERONI)


def _applicability_warnings(
    parameters: Mapping[str, float],
    system: SystemParameters,
    options: ApproxEGNOptions,
) -> list[str]:
    """Check the Eq. (3) assumptions and return recorded caveats."""

    warnings: list[str] = [
        "This is the attached paper's closed-form Approximate EGN correction, "
        "not a full EGN integral solver.",
        "The correction is applied to main.py's GN baseline without fitting.",
    ]
    rate_hz = float(parameters["rate_hz"])
    spacing_hz = float(parameters["spacing_hz"])
    if isinstance(system.channels, bool) or int(system.channels) != system.channels:
        raise ValueError("Approximate EGN의 channels는 정수여야 합니다.")
    if system.channels < 3:
        raise ValueError(
            "Eq. (3)은 동일 간격 WDM 시스템용입니다. channels는 3 이상으로 지정하세요."
        )
    if spacing_hz < rate_hz:
        message = (
            "채널 간격이 심볼율보다 작습니다. Eq. (3)의 비중첩 사각 스펙트럼 "
            "가정과 맞지 않습니다."
        )
        if options.strict_applicability:
            raise ValueError(message)
        warnings.append(message)
    if int(system.channels) % 2 == 0:
        message = (
            "짝수 채널 수에는 정확한 중앙 채널이 없습니다. "
            "H_((Nch-1)/2)의 반정수 연속값으로 pseudo-centre 채널을 근사했습니다."
        )
        if not options.allow_even_channel_center_approximation:
            raise ValueError(message)
        warnings.append(message)
    if float(parameters["alpha_power_per_km"]) <= 0.0:
        raise ValueError("Approximate EGN 보정식은 0보다 큰 광섬유 손실을 가정합니다.")
    return warnings


def approximate_egn_correction_eta_per_span(
    parameters: Mapping[str, float],
    system: SystemParameters,
    options: ApproxEGNOptions = ApproxEGNOptions(),
) -> Dict[str, Any]:
    """Return the Eq. (3) NLI-coefficient correction in W^-2 per span.

    Eq. (3) gives G_corr with N_s and P_ch**3. Integrating its flat PSD over
    the channel symbol rate and normalizing by N_s * P_ch**3 gives the
    per-span eta correction used here:

        eta_corr = (80/81) * Phi * gamma^2 * Leff^2
                   / (Rs * Delta_f * pi * |beta2| * Ls)
                   * [H_((Nch-1)/2) + Delta_f/Rs]
    """

    warnings = _applicability_warnings(parameters, system, options)
    canonical_modulation = _canonical_modulation(options.modulation_format)
    phi = MODULATION_PHI[canonical_modulation]
    beta_abs = abs(float(parameters["beta2_s2_per_km"]))
    if beta_abs == 0.0:
        raise ValueError("Approximate EGN 보정식은 |beta2| > 0을 필요로 합니다.")

    rate_hz = float(parameters["rate_hz"])
    spacing_hz = float(parameters["spacing_hz"])
    gamma_per_w_km = float(parameters["gamma_per_w_km"])
    effective_length_km = float(parameters["effective_length_km"])
    span_length_km = float(system.span_length_km)
    harmonic_order = (float(system.channels) - 1.0) / 2.0
    harmonic = harmonic_number(harmonic_order)

    eta_correction = (
        (80.0 / 81.0)
        * phi
        * gamma_per_w_km**2
        * effective_length_km**2
        / (rate_hz * spacing_hz * np.pi * beta_abs * span_length_km)
        * (harmonic + spacing_hz / rate_hz)
    )
    if not np.isfinite(eta_correction) or eta_correction < 0.0:
        raise ValueError("계산된 Approximate EGN 보정 eta가 유효하지 않습니다.")

    return {
        "eta_correction_per_span_w_inv2": float(eta_correction),
        "modulation_format": canonical_modulation,
        "modulation_phi": float(phi),
        "harmonic_order": float(harmonic_order),
        "harmonic_number": float(harmonic),
        "warnings": warnings,
    }


def approximate_egn_eta_per_span(
    fiber: FiberParameters = FiberParameters(),
    system: SystemParameters = SystemParameters(),
    gn_options: GNOptions = GNOptions(),
    egn_options: ApproxEGNOptions = ApproxEGNOptions(),
) -> Dict[str, Any]:
    """Calculate GN eta, the Eq. (3) correction eta, and corrected EGN eta.

    main.py uses its own explicitly recorded GN coefficient (8/27 by
    default). This function does not silently replace that coefficient. For a
    literature comparison, ensure that the chosen GN PSD/power convention is
    the same convention assumed by the reference result.
    """

    parameters = derive_parameters(fiber, system)
    eta_gn = gn_eta_per_span(parameters, system, gn_options)
    correction = approximate_egn_correction_eta_per_span(
        parameters, system, egn_options
    )
    eta_approximate_egn = eta_gn - correction["eta_correction_per_span_w_inv2"]
    warnings = list(correction["warnings"])

    if eta_approximate_egn <= 0.0:
        message = (
            "EGN 보정값이 GN eta 이상입니다. GN 기준식과 Eq. (3)의 PSD/전력 "
            "정의 또는 적용 조건을 확인하세요."
        )
        if egn_options.strict_applicability:
            raise ValueError(message)
        warnings.append(message + " 결과 eta는 0으로 제한했습니다.")
        eta_approximate_egn = 0.0

    correction_ratio = correction["eta_correction_per_span_w_inv2"] / eta_gn
    warnings.append(
        "GN baseline coefficient is retained from main.py "
        f"(nli_coefficient={gn_options.nli_coefficient:g}); it was not fitted."
    )
    return {
        "eta_gn_per_span_w_inv2": float(eta_gn),
        "eta_correction_per_span_w_inv2": float(
            correction["eta_correction_per_span_w_inv2"]
        ),
        "eta_approx_egn_per_span_w_inv2": float(eta_approximate_egn),
        "nli_eta_reduction_percent": float(100.0 * correction_ratio),
        "parameters": parameters,
        "modulation_format": correction["modulation_format"],
        "modulation_phi": correction["modulation_phi"],
        "harmonic_order": correction["harmonic_order"],
        "harmonic_number": correction["harmonic_number"],
        "warnings": warnings,
    }


def _column_as_float_array(table: Any, column: str) -> np.ndarray:
    """Extract a float column from a pandas DataFrame or mapping."""

    values = table[column]
    if hasattr(values, "to_numpy"):
        return values.to_numpy(dtype=float)
    return np.asarray(values, dtype=float)


def simulate_gn_and_approximate_egn(
    launch_dbm: Sequence[float] | np.ndarray | float,
    spans: int,
    *,
    fiber: FiberParameters = FiberParameters(),
    system: SystemParameters = SystemParameters(),
    gn_options: GNOptions = GNOptions(),
    egn_options: ApproxEGNOptions = ApproxEGNOptions(),
    include_transceiver_noise: bool = True,
) -> Dict[str, Any]:
    """Run the unchanged GN baseline and the closed-form Approximate EGN."""

    if isinstance(spans, bool) or int(spans) != spans or spans < 1:
        raise ValueError("spans는 1 이상의 정수여야 합니다.")
    spans = int(spans)
    eta_result = approximate_egn_eta_per_span(
        fiber=fiber,
        system=system,
        gn_options=gn_options,
        egn_options=egn_options,
    )
    parameters = eta_result["parameters"]
    gn = calculate_snr(
        launch_dbm,
        spans,
        eta_result["eta_gn_per_span_w_inv2"],
        parameters,
        system,
        include_transceiver_noise=include_transceiver_noise,
    )
    approximate_egn = calculate_snr(
        launch_dbm,
        spans,
        eta_result["eta_approx_egn_per_span_w_inv2"],
        parameters,
        system,
        include_transceiver_noise=include_transceiver_noise,
    )
    nli_reduction_percent = (
        100.0
        * (gn["nli_w"] - approximate_egn["nli_w"])
        / np.maximum(gn["nli_w"], np.finfo(float).tiny)
    )
    curve = records_to_dataframe(
        {
            "launch_dbm": gn["launch_dbm"],
            "spans": gn["spans"],
            "distance_km": gn["distance_km"],
            "signal_w": gn["signal_w"],
            "ase_w": gn["ase_w"],
            "trx_equivalent_noise_w": gn["trx_equivalent_noise_w"],
            "gn_nli_w": gn["nli_w"],
            "approx_egn_nli_w": approximate_egn["nli_w"],
            "nli_reduction_percent": nli_reduction_percent,
            "gn_snr_db": gn["snr_db"],
            "approx_egn_snr_db": approximate_egn["snr_db"],
            "snr_gain_db": approximate_egn["snr_db"] - gn["snr_db"],
            "eta_gn_per_span_w_inv2": np.full_like(
                gn["launch_dbm"], eta_result["eta_gn_per_span_w_inv2"]
            ),
            "eta_approx_egn_per_span_w_inv2": np.full_like(
                gn["launch_dbm"], eta_result["eta_approx_egn_per_span_w_inv2"]
            ),
        }
    )
    return {
        "curve": curve,
        **eta_result,
        "fiber": asdict(fiber),
        "system": asdict(system),
        "gn_options": asdict(gn_options),
        "approx_egn_options": asdict(egn_options),
        "metadata": {
            "reference": PAPER_REFERENCE,
            "equations": PAPER_EQUATIONS,
            "correction_type": "closed-form approximate EGN",
            "no_reference_fitting": True,
            "include_transceiver_noise": bool(include_transceiver_noise),
            "applicability_warnings": eta_result["warnings"],
        },
    }


def validate_approximate_egn_model(
    launch_dbm: Sequence[float] | np.ndarray | float,
    spans: int,
    *,
    fiber: FiberParameters = FiberParameters(),
    system: SystemParameters = SystemParameters(),
    gn_options: GNOptions = GNOptions(),
    egn_options: ApproxEGNOptions = ApproxEGNOptions(),
    include_transceiver_noise: bool = True,
    reference_snr_db: Optional[Sequence[float] | np.ndarray] = None,
    reference_launch_dbm: Optional[Sequence[float] | np.ndarray] = None,
) -> Dict[str, Any]:
    """Run App. EGN and optionally score it against an external SNR curve.

    Reference values are used only for error calculation. No parameter,
    launch power, or eta scale is optimized to make the curves agree.
    """

    result = simulate_gn_and_approximate_egn(
        launch_dbm,
        spans,
        fiber=fiber,
        system=system,
        gn_options=gn_options,
        egn_options=egn_options,
        include_transceiver_noise=include_transceiver_noise,
    )
    if reference_snr_db is not None:
        reference_x = (
            np.asarray(launch_dbm, dtype=float)
            if reference_launch_dbm is None
            else np.asarray(reference_launch_dbm, dtype=float)
        )
        curve = result["curve"]
        predicted_x = _column_as_float_array(curve, "launch_dbm")
        result["validation"] = compare_curves(
            reference_x,
            reference_snr_db,
            predicted_x,
            _column_as_float_array(curve, "approx_egn_snr_db"),
        )
        result["gn_baseline_validation"] = compare_curves(
            reference_x,
            reference_snr_db,
            predicted_x,
            _column_as_float_array(curve, "gn_snr_db"),
        )
        result["metadata"]["reference_use"] = (
            "Error metrics only; no fitting or optimization was performed."
        )
    return result


def _write_table_csv(table: Any, path: Path) -> None:
    """Save a DataFrame or mapping as CSV without requiring pandas."""

    if hasattr(table, "to_csv"):
        table.to_csv(path, index=False, encoding="utf-8-sig")
        return
    if not isinstance(table, Mapping):
        raise TypeError("CSV로 저장할 수 없는 결과 형식입니다.")
    keys = list(table)
    rows = zip(*[np.asarray(table[key]).tolist() for key in keys])
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(keys)
        writer.writerows(rows)


def _json_default(value: Any):
    if isinstance(value, np.generic):
        return value.item()
    if isinstance(value, np.ndarray):
        return value.tolist()
    raise TypeError(f"JSON으로 변환할 수 없는 형식: {type(value)}")


def save_approximate_egn_result(
    result: Mapping[str, Any],
    directory: str | Path,
    *,
    prefix: str = "approx_egn",
) -> Path:
    """Save the GN/App. EGN curve, metadata, and optional validation tables."""

    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    _write_table_csv(result["curve"], directory / f"{prefix}_curve.csv")

    metadata = {
        "fiber": result.get("fiber"),
        "system": result.get("system"),
        "gn_options": result.get("gn_options"),
        "approx_egn_options": result.get("approx_egn_options"),
        "parameters": result.get("parameters"),
        "eta_gn_per_span_w_inv2": result.get("eta_gn_per_span_w_inv2"),
        "eta_correction_per_span_w_inv2": result.get(
            "eta_correction_per_span_w_inv2"
        ),
        "eta_approx_egn_per_span_w_inv2": result.get(
            "eta_approx_egn_per_span_w_inv2"
        ),
        "nli_eta_reduction_percent": result.get("nli_eta_reduction_percent"),
        "metadata": result.get("metadata"),
        "python": platform.python_version(),
        "numpy": np.__version__,
    }
    with (directory / f"{prefix}_metadata.json").open("w", encoding="utf-8") as handle:
        json.dump(metadata, handle, indent=2, ensure_ascii=False, default=_json_default)

    if "validation" in result:
        _write_table_csv(
            result["validation"]["comparison"],
            directory / f"{prefix}_validation.csv",
        )
        _write_table_csv(
            result["gn_baseline_validation"]["comparison"],
            directory / f"{prefix}_gn_baseline_validation.csv",
        )
        with (directory / f"{prefix}_metrics.json").open(
            "w", encoding="utf-8"
        ) as handle:
            json.dump(
                {
                    "approx_egn": result["validation"]["metrics"],
                    "gn_baseline": result["gn_baseline_validation"]["metrics"],
                },
                handle,
                indent=2,
                ensure_ascii=False,
                default=_json_default,
            )
    return directory


def plot_gn_and_approximate_egn(
    result: Mapping[str, Any],
    path: Optional[str | Path] = None,
) -> None:
    """Plot the GN and approximate-EGN SNR curves."""

    try:
        import matplotlib.pyplot as plt
    except ImportError as exc:
        raise RuntimeError("그래프에는 matplotlib가 필요합니다.") from exc

    curve = result["curve"]
    x = _column_as_float_array(curve, "launch_dbm")
    gn_snr = _column_as_float_array(curve, "gn_snr_db")
    egn_snr = _column_as_float_array(curve, "approx_egn_snr_db")
    fig, ax = plt.subplots(figsize=(9, 5.5), layout="constrained")
    ax.plot(x, gn_snr, lw=2.0, label="main.py GN baseline")
    ax.plot(x, egn_snr, lw=2.0, label="Approximate EGN (Eq. 1 + Eq. 3)")
    if "validation" in result:
        comparison = result["validation"]["comparison"]
        ax.plot(
            _column_as_float_array(comparison, "launch_dbm"),
            _column_as_float_array(comparison, "reference_snr_db"),
            "o",
            ms=3.5,
            label="external reference",
        )
    ax.set_xlabel("Launch power per channel, total DP (dBm)")
    ax.set_ylabel("End-to-end SNR (dB)")
    ax.grid(alpha=0.3)
    ax.legend()
    if path is None:
        plt.show()
    else:
        fig.savefig(Path(path), dpi=180)
        plt.close(fig)


def run_cli(argv: Optional[Sequence[str]] = None) -> Dict[str, Any]:
    """Command-line entry point using main.py-compatible default inputs."""

    default_fiber = FiberParameters()
    default_system = SystemParameters()
    default_gn_options = GNOptions()
    parser = argparse.ArgumentParser(
        description=(
            "main.py GN baseline plus the Poggiolini et al. closed-form "
            "Approximate EGN correction"
        )
    )
    parser.add_argument("--spans", type=int, default=30)
    parser.add_argument("--launch-min", type=float, default=-10.0)
    parser.add_argument("--launch-max", type=float, default=10.0)
    parser.add_argument("--launch-step", type=float, default=0.05)
    parser.add_argument(
        "--modulation",
        choices=list(MODULATION_PHI),
        default="PM-64QAM",
        help="Paper Eq. (3) modulation factor",
    )
    parser.add_argument("--channels", type=int, default=default_system.channels)
    parser.add_argument(
        "--symbol-rate-gbd", type=float, default=default_system.symbol_rate_gbd
    )
    parser.add_argument("--spacing-ghz", type=float, default=default_system.spacing_ghz)
    parser.add_argument(
        "--span-length-km", type=float, default=default_system.span_length_km
    )
    parser.add_argument(
        "--noise-figure-db", type=float, default=default_system.noise_figure_db
    )
    parser.add_argument("--trx-snr-db", type=float, default=default_system.transceiver_snr_db)
    parser.add_argument("--fiber-name", default=default_fiber.name)
    parser.add_argument(
        "--attenuation-db-per-km",
        type=float,
        default=default_fiber.attenuation_db_per_km,
    )
    parser.add_argument(
        "--dispersion-ps-nm-km",
        type=float,
        default=default_fiber.dispersion_ps_nm_km,
    )
    parser.add_argument(
        "--effective-area-um2", type=float, default=default_fiber.effective_area_um2
    )
    parser.add_argument("--n2-m2-w", type=float, default=default_fiber.n2_m2_w)
    parser.add_argument(
        "--gamma-per-w-km",
        type=float,
        default=None,
        help="Specify gamma directly; otherwise derive it from n2 and Aeff.",
    )
    parser.add_argument(
        "--gn-coefficient",
        type=float,
        default=default_gn_options.nli_coefficient,
        help="Retained main.py GN coefficient; no automatic fitting is done.",
    )
    parser.add_argument(
        "--exclude-transceiver-noise",
        action="store_true",
        help="Use only ASE + NLI in the SNR calculation.",
    )
    parser.add_argument(
        "--reject-even-channels",
        action="store_true",
        help="Reject, instead of approximate, an even-channel pseudo-centre CUT.",
    )
    parser.add_argument(
        "--allow-out-of-scope",
        action="store_true",
        help="Warn rather than reject a spacing smaller than the symbol rate.",
    )
    parser.add_argument(
        "--reference-csv",
        type=str,
        default=None,
        help="Optional CSV with launch_dbm and snr_db for error metrics only.",
    )
    parser.add_argument("--save-dir", type=str, default="approx_egn_results")
    parser.add_argument("--plot", action="store_true")
    args = parser.parse_args(argv)

    fiber = FiberParameters(
        name=args.fiber_name,
        attenuation_db_per_km=args.attenuation_db_per_km,
        effective_area_um2=args.effective_area_um2,
        dispersion_ps_nm_km=args.dispersion_ps_nm_km,
        n2_m2_w=args.n2_m2_w,
        gamma_per_w_km=args.gamma_per_w_km,
    )
    system = SystemParameters(
        channels=args.channels,
        symbol_rate_gbd=args.symbol_rate_gbd,
        spacing_ghz=args.spacing_ghz,
        span_length_km=args.span_length_km,
        noise_figure_db=args.noise_figure_db,
        transceiver_snr_db=args.trx_snr_db,
    )
    gn_options = GNOptions(
        nli_coefficient=args.gn_coefficient,
        finite_effective_length=True,
        model_name="gn_finite_main_py_baseline",
    )
    egn_options = ApproxEGNOptions(
        modulation_format=args.modulation,
        strict_applicability=not args.allow_out_of_scope,
        allow_even_channel_center_approximation=not args.reject_even_channels,
    )
    if args.reference_csv is None:
        launch_dbm = make_launch_grid(
            args.launch_min, args.launch_max, args.launch_step
        )
        reference_launch_dbm = reference_snr_db = None
    else:
        reference_launch_dbm, reference_snr_db = read_reference_csv(
            args.reference_csv, spans=args.spans
        )
        launch_dbm = reference_launch_dbm

    result = validate_approximate_egn_model(
        launch_dbm,
        args.spans,
        fiber=fiber,
        system=system,
        gn_options=gn_options,
        egn_options=egn_options,
        include_transceiver_noise=not args.exclude_transceiver_noise,
        reference_snr_db=reference_snr_db,
        reference_launch_dbm=reference_launch_dbm,
    )
    output_dir = save_approximate_egn_result(result, args.save_dir)
    if args.plot:
        plot_gn_and_approximate_egn(result, output_dir / "gn_vs_approx_egn.png")

    print(f"Saved GN/App. EGN results to: {output_dir}")
    print(f"GN eta per span: {result['eta_gn_per_span_w_inv2']:.6e} W^-2")
    print(
        "App. EGN correction eta per span: "
        f"{result['eta_correction_per_span_w_inv2']:.6e} W^-2"
    )
    print(
        "App. EGN eta per span: "
        f"{result['eta_approx_egn_per_span_w_inv2']:.6e} W^-2 "
        f"({result['nli_eta_reduction_percent']:.2f}% lower than GN)"
    )
    if "validation" in result:
        print(f"App. EGN validation: {result['validation']['metrics']}")
        print(f"GN baseline validation: {result['gn_baseline_validation']['metrics']}")
    for warning in result["warnings"]:
        print(f"Note: {warning}")
    return result


def main(argv: Optional[Sequence[str]] = None) -> Dict[str, Any]:
    """CLI-compatible entry point."""

    return run_cli(argv)


if __name__ == "__main__":
    main()
