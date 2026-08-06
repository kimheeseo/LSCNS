"""Reproduce and validate Fig. 2 of Sohanpal et al. (arXiv:2606.17942).

The script produces two deliberately separate results:

1. ``fig2_reproduction.svg`` redraws the exact data distributed with the
   paper's arXiv source.
2. ``fig2_model_validation.svg`` compares those data with the parameter-based
   GN calculation in this repository and with an equivalent Eq. (1) fit.

The HCF parameters are stated explicitly in the paper. The exact flat C-band
SMF loss/dispersion/nonlinearity profile used for Fig. 2 is not tabulated, so
the independent SMF calculation uses the editable parent-repository baseline.
The equivalent fit is a diagnostic, not an independently predicted result.

Run from the gn-model-gsnr project root:

    python hcf_vs_2core/fig2_reproduction/reproduce_fig2.py
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

import numpy as np


PLANCK_J_S = 6.62607015e-34
C_NM_PER_S = 2.99792458e17
DB_PER_NEPER = 4.343


@dataclass(frozen=True)
class LinkCase:
    name: str
    short_name: str
    total_length_km: float
    span_count: int
    attenuation_db_per_km: float
    dispersion_s_per_nm_km: float
    nonlinearity_per_w_km: float
    noise_figure_db: float
    distributed_interference_db_per_km: float | None


HCF = LinkCase(
    name="HCF: 1 x 200 km",
    short_name="HCF",
    total_length_km=200.0,
    span_count=1,
    attenuation_db_per_km=0.075,
    dispersion_s_per_nm_km=3.2e-12,
    nonlinearity_per_w_km=5.0e-4,
    noise_figure_db=5.0,
    distributed_interference_db_per_km=-52.0,
)

# Low-loss large-effective-area SMF baseline used by the parent repository.
# It is exposed here because the paper does not tabulate the corresponding
# flat C-band SMF parameters used to generate Fig. 2.
SMF = LinkCase(
    name="Low-loss SMF: 3 x 67 km",
    short_name="SMF",
    total_length_km=200.0,
    span_count=3,
    attenuation_db_per_km=0.15,
    dispersion_s_per_nm_km=-21.0e-12,
    nonlinearity_per_w_km=0.81,
    noise_figure_db=5.0,
    distributed_interference_db_per_km=None,
)


CENTRE_WAVELENGTH_NM = 1550.0
SYMBOL_RATE_BAUD = 140e9
CHANNEL_SPACING_HZ = 150e9
CHANNEL_COUNT = 29
TRANSCEIVER_SNR_DB = 20.0

PLOT_MIN_POWER_DBM = -40.0
PLOT_MAX_POWER_DBM = 50.0
# The model grid covers every point in both source-data files; the figures crop
# it to the paper's published -40 to 50 dBm viewing window.
MODEL_POWER_DBM = np.linspace(-60.0, 60.0, 1201)
FIT_MIN_POWER_DBM = -20.0
FIT_MAX_POWER_DBM = 40.0


def db_to_linear(value_db: float | np.ndarray) -> np.ndarray:
    return np.power(10.0, np.asarray(value_db, dtype=float) / 10.0)


def dbm_to_watts(value_dbm: float | np.ndarray) -> np.ndarray:
    return np.power(10.0, (np.asarray(value_dbm, dtype=float) - 30.0) / 10.0)


def closed_form_nli_coefficient(case: LinkCase) -> float:
    """Return the parent model's per-span inverse-SNR NLI coefficient."""
    span_length_km = case.total_length_km / case.span_count
    alpha_linear_per_km = case.attenuation_db_per_km / DB_PER_NEPER
    effective_length_km = (
        1.0 - np.exp(-alpha_linear_per_km * span_length_km)
    ) / alpha_linear_per_km
    asymptotic_effective_length_km = 1.0 / alpha_linear_per_km
    beta_2_s2_per_km = (
        -case.dispersion_s_per_nm_km
        * CENTRE_WAVELENGTH_NM**2
        / (2.0 * np.pi * C_NM_PER_S)
    )
    abs_beta_2 = abs(beta_2_s2_per_km)
    asinh_argument = (
        np.pi**2
        / 2.0
        * abs_beta_2
        * asymptotic_effective_length_km
        * SYMBOL_RATE_BAUD**2
        * CHANNEL_COUNT ** (2.0 * SYMBOL_RATE_BAUD / CHANNEL_SPACING_HZ)
    )
    return float(
        8.0
        / 27.0
        * case.nonlinearity_per_w_km**2
        * effective_length_km**2
        * np.arcsinh(asinh_argument)
        / (
            np.pi
            * abs_beta_2
            * asymptotic_effective_length_km
            * SYMBOL_RATE_BAUD**2
        )
    )


def fixed_inverse_snr(case: LinkCase) -> float:
    value = 1.0 / float(db_to_linear(TRANSCEIVER_SNR_DB))
    if case.distributed_interference_db_per_km is not None:
        value += (
            float(db_to_linear(case.distributed_interference_db_per_km))
            * case.total_length_km
        )
    return value


def physical_noise_coefficients(case: LinkCase) -> tuple[float, float]:
    """Return total ASE power A and total eta for 1/SNR=A/P+eta*P^2+c."""
    span_length_km = case.total_length_km / case.span_count
    span_gain_linear = float(
        db_to_linear(case.attenuation_db_per_km * span_length_km)
    )
    optical_frequency_hz = C_NM_PER_S / CENTRE_WAVELENGTH_NM
    ase_per_span_w = (
        PLANCK_J_S
        * optical_frequency_hz
        * float(db_to_linear(case.noise_figure_db))
        * SYMBOL_RATE_BAUD
        * (span_gain_linear - 1.0)
    )
    total_ase_w = case.span_count * ase_per_span_w
    total_eta_per_w2 = case.span_count * closed_form_nli_coefficient(case)
    return total_ase_w, total_eta_per_w2


def curve_from_coefficients(
    case: LinkCase,
    power_dbm: np.ndarray,
    ase_w: float,
    eta_per_w2: float,
) -> dict[str, np.ndarray]:
    channel_power_w = dbm_to_watts(power_dbm)
    inverse_snr = (
        ase_w / channel_power_w
        + eta_per_w2 * channel_power_w**2
        + fixed_inverse_snr(case)
    )
    system_snr = 1.0 / inverse_snr
    throughput_tbps = (
        2.0
        * SYMBOL_RATE_BAUD
        * CHANNEL_COUNT
        * np.log2(1.0 + system_snr)
        / 1e12
    )
    return {
        "power_dbm": power_dbm,
        "system_snr_linear": system_snr,
        "throughput_tbps": throughput_tbps,
    }


def calculate_physical_case(case: LinkCase) -> dict[str, np.ndarray]:
    ase_w, eta_per_w2 = physical_noise_coefficients(case)
    return curve_from_coefficients(case, MODEL_POWER_DBM, ase_w, eta_per_w2)


def read_reference(path: Path) -> dict[str, np.ndarray]:
    data = np.genfromtxt(path, delimiter="\t", names=True)
    return {
        "power_dbm": np.asarray(data["lpch"], dtype=float),
        "throughput_tbps": np.asarray(data["capc"], dtype=float),
    }


def throughput_to_equivalent_snr(throughput_tbps: np.ndarray) -> np.ndarray:
    spectral_efficiency = (
        throughput_tbps * 1e12 / (2.0 * SYMBOL_RATE_BAUD * CHANNEL_COUNT)
    )
    return np.exp2(spectral_efficiency) - 1.0


def fit_equivalent_coefficients(
    case: LinkCase,
    reference: dict[str, np.ndarray],
) -> tuple[float, float]:
    """Fit the two free coefficients of Eq. (1) on the visible central range."""
    power_dbm = reference["power_dbm"]
    power_w = dbm_to_watts(power_dbm)
    snr = throughput_to_equivalent_snr(reference["throughput_tbps"])
    target = 1.0 / snr - fixed_inverse_snr(case)
    fit_mask = (
        (power_dbm >= FIT_MIN_POWER_DBM)
        & (power_dbm <= FIT_MAX_POWER_DBM)
        & np.isfinite(target)
        & (target > 0.0)
    )
    design = np.column_stack((1.0 / power_w[fit_mask], power_w[fit_mask] ** 2))
    coefficients, _, _, _ = np.linalg.lstsq(design, target[fit_mask], rcond=None)
    return float(coefficients[0]), float(coefficients[1])


def peak(curve: dict[str, np.ndarray]) -> tuple[float, float]:
    index = int(np.nanargmax(curve["throughput_tbps"]))
    return (
        float(curve["power_dbm"][index]),
        float(curve["throughput_tbps"][index]),
    )


def reference_peak(reference: dict[str, np.ndarray]) -> tuple[float, float]:
    index = int(np.nanargmax(reference["throughput_tbps"]))
    return (
        float(reference["power_dbm"][index]),
        float(reference["throughput_tbps"][index]),
    )


def values_at_reference_points(
    curve: dict[str, np.ndarray],
    reference: dict[str, np.ndarray],
) -> np.ndarray:
    return np.interp(
        reference["power_dbm"],
        curve["power_dbm"],
        curve["throughput_tbps"],
    )


def visible_rmse(
    reference: dict[str, np.ndarray],
    calculated_values: np.ndarray,
) -> float:
    mask = (
        (reference["power_dbm"] >= PLOT_MIN_POWER_DBM)
        & (reference["power_dbm"] <= PLOT_MAX_POWER_DBM)
    )
    residual = calculated_values[mask] - reference["throughput_tbps"][mask]
    return float(np.sqrt(np.mean(residual**2)))


def write_comparison_csv(
    output_path: Path,
    cases: list[
        tuple[
            LinkCase,
            dict[str, np.ndarray],
            dict[str, np.ndarray],
            dict[str, np.ndarray],
        ]
    ],
) -> None:
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "case",
                "launch_power_per_channel_dbm",
                "paper_reference_throughput_tbps",
                "physical_model_throughput_tbps",
                "equivalent_eq1_fit_throughput_tbps",
                "physical_residual_tbps",
                "equivalent_fit_residual_tbps",
            ]
        )
        for case, reference, physical, fitted in cases:
            physical_values = values_at_reference_points(physical, reference)
            fitted_values = values_at_reference_points(fitted, reference)
            for index, power_dbm in enumerate(reference["power_dbm"]):
                reference_value = reference["throughput_tbps"][index]
                writer.writerow(
                    [
                        case.short_name,
                        f"{power_dbm:.1f}",
                        f"{reference_value:.12g}",
                        f"{physical_values[index]:.12g}",
                        f"{fitted_values[index]:.12g}",
                        f"{physical_values[index] - reference_value:.12g}",
                        f"{fitted_values[index] - reference_value:.12g}",
                    ]
                )


def write_metrics_csv(
    output_path: Path,
    cases: list[
        tuple[
            LinkCase,
            dict[str, np.ndarray],
            dict[str, np.ndarray],
            dict[str, np.ndarray],
            float,
            float,
        ]
    ],
) -> None:
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "case",
                "reference_peak_power_dbm_per_channel",
                "reference_peak_throughput_tbps",
                "physical_peak_power_dbm_per_channel",
                "physical_peak_throughput_tbps",
                "physical_rmse_tbps_visible_range",
                "equivalent_fit_ase_w",
                "equivalent_fit_eta_per_w2",
                "equivalent_fit_peak_power_dbm_per_channel",
                "equivalent_fit_peak_throughput_tbps",
                "equivalent_fit_rmse_tbps_visible_range",
            ]
        )
        for case, reference, physical, fitted, fit_ase, fit_eta in cases:
            ref_power, ref_capacity = reference_peak(reference)
            physical_power, physical_capacity = peak(physical)
            fit_power, fit_capacity = peak(fitted)
            physical_values = values_at_reference_points(physical, reference)
            fitted_values = values_at_reference_points(fitted, reference)
            writer.writerow(
                [
                    case.short_name,
                    f"{ref_power:.3f}",
                    f"{ref_capacity:.12f}",
                    f"{physical_power:.3f}",
                    f"{physical_capacity:.12f}",
                    f"{visible_rmse(reference, physical_values):.12f}",
                    f"{fit_ase:.12e}",
                    f"{fit_eta:.12e}",
                    f"{fit_power:.3f}",
                    f"{fit_capacity:.12f}",
                    f"{visible_rmse(reference, fitted_values):.12f}",
                ]
            )


def configure_axis(ax: object) -> None:
    ax.set_xlim(PLOT_MIN_POWER_DBM, PLOT_MAX_POWER_DBM)
    ax.set_ylim(1.0, 80.0)
    ax.set_yscale("log")
    ax.set_xlabel("Launch power per channel (dBm)")
    ax.set_ylabel("Throughput (Tb/s)")
    ax.set_yticks([1.0, 10.0, 60.0], labels=["1", "10", "60"])
    ax.grid(True, which="major", alpha=0.35)


def write_reference_figure(
    output_path: Path,
    hcf_reference: dict[str, np.ndarray],
    smf_reference: dict[str, np.ndarray],
) -> None:
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(7.4, 4.4))
    ax.plot(
        hcf_reference["power_dbm"],
        hcf_reference["throughput_tbps"],
        color="#0072B2",
        linewidth=2.3,
        label="HCF",
    )
    ax.plot(
        smf_reference["power_dbm"],
        smf_reference["throughput_tbps"],
        color="#D55E00",
        linewidth=2.3,
        label="SMF",
    )
    for reference, color in (
        (hcf_reference, "#0072B2"),
        (smf_reference, "#D55E00"),
    ):
        power_dbm, throughput_tbps = reference_peak(reference)
        ax.scatter(power_dbm, throughput_tbps, marker="s", color=color, s=38, zorder=4)
        ax.annotate(
            r"$T_{tot}^{max}$",
            (power_dbm, throughput_tbps),
            xytext=(0, -16),
            textcoords="offset points",
            ha="center",
            fontsize=9,
        )
    configure_axis(ax)
    ax.set_title("Fig. 2 reproduction from the paper's source data")
    ax.legend(loc="lower right")
    fig.tight_layout()
    fig.savefig(output_path, format="svg")
    plt.close(fig)


def write_validation_figure(
    output_path: Path,
    cases: list[
        tuple[
            LinkCase,
            dict[str, np.ndarray],
            dict[str, np.ndarray],
            dict[str, np.ndarray],
        ]
    ],
) -> None:
    import matplotlib.pyplot as plt

    fig, axes = plt.subplots(1, 2, figsize=(11.0, 4.3), sharey=True)
    colors = {"HCF": "#0072B2", "SMF": "#D55E00"}
    for ax, (case, reference, physical, fitted) in zip(axes, cases):
        color = colors[case.short_name]
        ax.plot(
            reference["power_dbm"],
            reference["throughput_tbps"],
            color=color,
            linewidth=2.5,
            label="Paper reference",
        )
        ax.plot(
            physical["power_dbm"],
            physical["throughput_tbps"],
            color="black",
            linestyle="--",
            linewidth=1.8,
            label="Physical-input model",
        )
        ax.plot(
            fitted["power_dbm"],
            fitted["throughput_tbps"],
            color="#009E73",
            linestyle=":",
            linewidth=2.1,
            label="Equivalent Eq. (1) fit",
        )
        configure_axis(ax)
        ax.set_title(case.name)
        ax.legend(loc="lower right", fontsize=8)
    axes[1].set_ylabel("")
    fig.suptitle("Fig. 2 model validation: disclosed inputs vs calibrated fit", y=1.01)
    fig.tight_layout()
    fig.savefig(output_path, format="svg", bbox_inches="tight")
    plt.close(fig)


def print_summary(
    case: LinkCase,
    reference: dict[str, np.ndarray],
    physical: dict[str, np.ndarray],
    fitted: dict[str, np.ndarray],
) -> None:
    ref_power, ref_capacity = reference_peak(reference)
    physical_power, physical_capacity = peak(physical)
    fit_power, fit_capacity = peak(fitted)
    physical_values = values_at_reference_points(physical, reference)
    fitted_values = values_at_reference_points(fitted, reference)
    print(
        f"{case.short_name}: paper {ref_capacity:.6f} Tb/s at {ref_power:.1f} dBm/ch; "
        f"physical model {physical_capacity:.6f} Tb/s at {physical_power:.1f} dBm/ch "
        f"(visible RMSE {visible_rmse(reference, physical_values):.6f} Tb/s); "
        f"Eq. (1) fit {fit_capacity:.6f} Tb/s at {fit_power:.1f} dBm/ch "
        f"(visible RMSE {visible_rmse(reference, fitted_values):.6f} Tb/s)"
    )


def main() -> None:
    output_dir = Path(__file__).resolve().parent
    hcf_reference = read_reference(output_dir / "reference_fig2_hcf.tsv")
    smf_reference = read_reference(output_dir / "reference_fig2_smf.tsv")

    hcf_physical = calculate_physical_case(HCF)
    smf_physical = calculate_physical_case(SMF)

    hcf_fit_ase, hcf_fit_eta = fit_equivalent_coefficients(HCF, hcf_reference)
    smf_fit_ase, smf_fit_eta = fit_equivalent_coefficients(SMF, smf_reference)
    hcf_fitted = curve_from_coefficients(
        HCF, MODEL_POWER_DBM, hcf_fit_ase, hcf_fit_eta
    )
    smf_fitted = curve_from_coefficients(
        SMF, MODEL_POWER_DBM, smf_fit_ase, smf_fit_eta
    )

    comparison_cases = [
        (HCF, hcf_reference, hcf_physical, hcf_fitted),
        (SMF, smf_reference, smf_physical, smf_fitted),
    ]
    metric_cases = [
        (HCF, hcf_reference, hcf_physical, hcf_fitted, hcf_fit_ase, hcf_fit_eta),
        (SMF, smf_reference, smf_physical, smf_fitted, smf_fit_ase, smf_fit_eta),
    ]

    write_reference_figure(
        output_dir / "fig2_reproduction.svg", hcf_reference, smf_reference
    )
    write_validation_figure(output_dir / "fig2_model_validation.svg", comparison_cases)
    write_comparison_csv(output_dir / "fig2_reproduction.csv", comparison_cases)
    write_metrics_csv(output_dir / "fig2_metrics.csv", metric_cases)

    print_summary(HCF, hcf_reference, hcf_physical, hcf_fitted)
    print_summary(SMF, smf_reference, smf_physical, smf_fitted)


if __name__ == "__main__":
    main()
