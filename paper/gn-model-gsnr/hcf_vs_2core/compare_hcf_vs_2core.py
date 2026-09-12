"""Compare a single-core HCF system with a two-core MCF system.

This script reuses ``gn_model_gsnr.optimized_capacity_extended`` without changing
the core GN/GSNR implementation.  The HCF inter-modal interference (IMI)
coefficient is mapped to the model's distributed additive-crosstalk input because
the current package has no separate IMI term.

Run from the ``gn-model-gsnr`` project root:

    python hcf_vs_2core/compare_hcf_vs_2core.py

The script writes ``results.csv`` and ``capacity_vs_span.svg`` beside itself.
"""

from __future__ import annotations

import csv
import sys
from dataclasses import dataclass
from pathlib import Path

import numpy as np

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from gn_model_gsnr import optimized_capacity_extended


@dataclass(frozen=True)
class FibreCase:
    name: str
    attenuation_db_per_km: float
    nonlinearity_per_w_km: float
    dispersion_s_per_nm_km: float
    distributed_interference_db_per_km: float
    fifo_loss_db: float
    noise_figure_db: float
    n_cores: int


HCF = FibreCase(
    name="HCF (1 core)",
    attenuation_db_per_km=0.075,
    nonlinearity_per_w_km=5.0e-4,
    dispersion_s_per_nm_km=3.2e-12,
    distributed_interference_db_per_km=-52.0,
    fifo_loss_db=0.0,
    noise_figure_db=5.0,
    n_cores=1,
)

MCF_2CORE = FibreCase(
    name="MCF (2 core)",
    attenuation_db_per_km=0.15,
    nonlinearity_per_w_km=0.81,
    dispersion_s_per_nm_km=-21.0e-12,
    distributed_interference_db_per_km=-80.0,
    fifo_loss_db=0.3,
    noise_figure_db=4.5,
    n_cores=2,
)


TOTAL_LENGTH_KM = 6000.0
FIBRE_PAIRS = 12
PFE_VOLTAGE_V = 18_000.0
CONDUCTOR_RESISTANCE_OHM_PER_KM = 1.0
HOUSEKEEPING_FRACTION = 0.10
AMPLIFIER_PCE = 0.025
CENTRE_WAVELENGTH_NM = 1550.0
OCCUPIED_BANDWIDTH_HZ = 4.3e12
SYMBOL_RATE_BAUD = 100e9
CHANNEL_SPACING_HZ = 112.5e9
TRANSCEIVER_SNR_DB = 20.0
LAUNCH_POWER_CAP_DBM = 18.0
NONLINEAR_BACKOFF_DB = 2.0

SPAN_LENGTHS_KM = np.arange(50.0, 251.0, 5.0)
# Extending well beyond the 18 dBm hardware cap is necessary to locate the HCF
# nonlinear threshold instead of mistaking the end of the grid for that threshold.
LAUNCH_POWERS_W = np.geomspace(1e-4, 100.0, 2400)


def combine_with_transceiver_snr(link_snr: np.ndarray) -> np.ndarray:
    """Combine optical-link GSNR and transceiver SNR as independent noises."""
    trx_snr = 10.0 ** (TRANSCEIVER_SNR_DB / 10.0)
    return 1.0 / (1.0 / link_snr + 1.0 / trx_snr)


def simulate(case: FibreCase) -> dict[str, np.ndarray]:
    efficiency = np.full(
        (len(SPAN_LENGTHS_KM), len(LAUNCH_POWERS_W)), AMPLIFIER_PCE
    )
    result = optimized_capacity_extended(
        SPAN_LENGTHS_KM,
        TOTAL_LENGTH_KM,
        case.noise_figure_db,
        PFE_VOLTAGE_V,
        CONDUCTOR_RESISTANCE_OHM_PER_KM,
        efficiency,
        HOUSEKEEPING_FRACTION,
        FIBRE_PAIRS,
        case.nonlinearity_per_w_km,
        case.distributed_interference_db_per_km,
        case.attenuation_db_per_km,
        case.fifo_loss_db,
        case.dispersion_s_per_nm_km,
        SYMBOL_RATE_BAUD,
        CHANNEL_SPACING_HZ,
        LAUNCH_POWERS_W,
        CENTRE_WAVELENGTH_NM,
        OCCUPIED_BANDWIDTH_HZ,
        float(SPAN_LENGTHS_KM[0]),
        n_cores=case.n_cores,
        count_mode="physical",
        gamma_xt_unit="dB_per_km",
        nonlinear_backoff_db=NONLINEAR_BACKOFF_DB,
        launch_power_cap_dbm=LAUNCH_POWER_CAP_DBM,
        return_details=True,
    )
    system_gsnr = combine_with_transceiver_snr(result.gsnr_by_span_linear)
    channel_count = int(np.floor(OCCUPIED_BANDWIDTH_HZ / CHANNEL_SPACING_HZ))
    capacity_bps = (
        FIBRE_PAIRS
        * case.n_cores
        * channel_count
        * SYMBOL_RATE_BAUD
        * np.log2(1.0 + system_gsnr)
    )
    return {
        "link_gsnr_linear": result.gsnr_by_span_linear,
        "system_gsnr_linear": system_gsnr,
        "capacity_bps": capacity_bps,
        "launch_power_w": result.launch_power_by_span_w,
        "repeater_count": result.repeater_count,
        "feasible": result.feasible,
    }


def safe_db(values: np.ndarray) -> np.ndarray:
    return 10.0 * np.log10(values)


def safe_dbm(values_w: np.ndarray) -> np.ndarray:
    return 10.0 * np.log10(values_w) + 30.0


def optimum_index(values: np.ndarray) -> int:
    return int(np.nanargmax(values))


def write_csv(output_path: Path, hcf: dict[str, np.ndarray], mcf: dict[str, np.ndarray]) -> None:
    fields = [
        "span_km",
        "hcf_capacity_tbps",
        "mcf_2core_capacity_tbps",
        "hcf_system_gsnr_db",
        "mcf_2core_system_gsnr_db",
        "hcf_launch_power_dbm",
        "mcf_2core_launch_power_dbm",
        "hcf_repeaters",
        "mcf_2core_repeaters",
        "hcf_feasible",
        "mcf_2core_feasible",
    ]
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for index, span in enumerate(SPAN_LENGTHS_KM):
            writer.writerow(
                {
                    "span_km": f"{span:.0f}",
                    "hcf_capacity_tbps": f"{hcf['capacity_bps'][index] / 1e12:.6f}",
                    "mcf_2core_capacity_tbps": f"{mcf['capacity_bps'][index] / 1e12:.6f}",
                    "hcf_system_gsnr_db": f"{safe_db(hcf['system_gsnr_linear'])[index]:.6f}",
                    "mcf_2core_system_gsnr_db": f"{safe_db(mcf['system_gsnr_linear'])[index]:.6f}",
                    "hcf_launch_power_dbm": f"{safe_dbm(hcf['launch_power_w'])[index]:.6f}",
                    "mcf_2core_launch_power_dbm": f"{safe_dbm(mcf['launch_power_w'])[index]:.6f}",
                    "hcf_repeaters": int(hcf["repeater_count"][index]),
                    "mcf_2core_repeaters": int(mcf["repeater_count"][index]),
                    "hcf_feasible": bool(hcf["feasible"][index]),
                    "mcf_2core_feasible": bool(mcf["feasible"][index]),
                }
            )


def write_plot(output_path: Path, hcf: dict[str, np.ndarray], mcf: dict[str, np.ndarray]) -> None:
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(8.5, 5.2))
    ax.plot(
        SPAN_LENGTHS_KM,
        hcf["capacity_bps"] / 1e12,
        linewidth=2.2,
        label=HCF.name,
    )
    ax.plot(
        SPAN_LENGTHS_KM,
        mcf["capacity_bps"] / 1e12,
        linewidth=2.2,
        label=MCF_2CORE.name,
    )
    for case, data in ((HCF, hcf), (MCF_2CORE, mcf)):
        idx = optimum_index(data["capacity_bps"])
        ax.scatter(
            SPAN_LENGTHS_KM[idx], data["capacity_bps"][idx] / 1e12, s=45
        )
        ax.annotate(
            f"{case.name}: {SPAN_LENGTHS_KM[idx]:.0f} km",
            (SPAN_LENGTHS_KM[idx], data["capacity_bps"][idx] / 1e12),
            xytext=(7, 7),
            textcoords="offset points",
            fontsize=8,
        )
    ax.set_xlabel("Candidate span length (km)")
    ax.set_ylabel("One-direction aggregate capacity (Tb/s)")
    ax.set_title("6000 km submarine link: HCF versus two-core MCF")
    ax.grid(True, alpha=0.3)
    ax.legend()
    fig.tight_layout()
    fig.savefig(output_path, format="svg")
    plt.close(fig)


def print_summary(hcf: dict[str, np.ndarray], mcf: dict[str, np.ndarray]) -> None:
    hcf_idx = optimum_index(hcf["capacity_bps"])
    mcf_idx = optimum_index(mcf["capacity_bps"])
    hcf_capacity = hcf["capacity_bps"][hcf_idx]
    mcf_capacity = mcf["capacity_bps"][mcf_idx]
    print(
        f"HCF optimum: {SPAN_LENGTHS_KM[hcf_idx]:.0f} km, "
        f"{hcf_capacity / 1e12:.3f} Tb/s, "
        f"GSNR {safe_db(hcf['system_gsnr_linear'])[hcf_idx]:.2f} dB, "
        f"launch {safe_dbm(hcf['launch_power_w'])[hcf_idx]:.2f} dBm"
    )
    print(
        f"2-core MCF optimum: {SPAN_LENGTHS_KM[mcf_idx]:.0f} km, "
        f"{mcf_capacity / 1e12:.3f} Tb/s, "
        f"GSNR {safe_db(mcf['system_gsnr_linear'])[mcf_idx]:.2f} dB, "
        f"launch {safe_dbm(mcf['launch_power_w'])[mcf_idx]:.2f} dBm"
    )
    print(f"Peak-capacity HCF/MCF ratio: {hcf_capacity / mcf_capacity:.3f}")


def main() -> None:
    output_dir = Path(__file__).resolve().parent
    hcf = simulate(HCF)
    mcf = simulate(MCF_2CORE)
    write_csv(output_dir / "results.csv", hcf, mcf)
    write_plot(output_dir / "capacity_vs_span.svg", hcf, mcf)
    print_summary(hcf, mcf)


if __name__ == "__main__":
    main()
