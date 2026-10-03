"""No-fit HCF2 loss validation used by the executed Colab report.

The Petrovich HCF2 numbers enter only after all model predictions have been
calculated.  They are benchmark values, never optimizer targets.
"""

from __future__ import annotations

import json
import platform
import sys
import time
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import scipy

from hcf_concentric_tmm_20260908 import (
    bache_bouncing_ray_loss_db_km,
    bouncing_ray_te_loss_db_km,
    build_models,
    multilayer_bouncing_ray_loss_db_km,
    silica_index_sellmeier,
    solve_fundamental_mode,
    solve_models,
)


OUTPUT_DIR = Path(__file__).resolve().parent / "results"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Petrovich HCF2 geometry.  Diameter intervals are measured ranges published
# for the 15-km fibre; the wall thickness is reported as approximately 500 nm.
WAVELENGTHS_NM = (1310.0, 1550.0)
# Keep the user's previous benchmark condition exactly.  The corresponding
# 29.50-um core diameter lies inside the paper's 29.1-29.6-um measured range.
CORE_RADIUS_UM = 14.75
WALL_THICKNESS_UM = 0.50
TUBE_COUNT = 5
LARGE_DIAMETER_RANGE_UM = (30.4, 31.7)
MIDDLE_DIAMETER_RANGE_UM = (22.7, 24.8)
SMALL_DIAMETER_RANGE_UM = (7.0, 8.4)

# These are held out from every solver and introduced only in benchmark code.
PETROVICH_MEASURED_TOTAL_DB_KM = {1310.0: 0.128, 1550.0: 0.091}
TARGET_USED_FOR_FITTING = False


def midpoint(bounds: tuple[float, float]) -> float:
    return float(sum(bounds) / 2.0)


LARGE_DIAMETER_UM = midpoint(LARGE_DIAMETER_RANGE_UM)
MIDDLE_DIAMETER_UM = midpoint(MIDDLE_DIAMETER_RANGE_UM)
SMALL_DIAMETER_UM = midpoint(SMALL_DIAMETER_RANGE_UM)

# Internally tangent one-dimensional radial map.  This is an explicit
# surrogate assumption, not a measured concentric air-gap thickness.
GAP1_UM = (LARGE_DIAMETER_UM - 2 * WALL_THICKNESS_UM) - MIDDLE_DIAMETER_UM
GAP2_UM = (MIDDLE_DIAMETER_UM - 2 * WALL_THICKNESS_UM) - SMALL_DIAMETER_UM


def signed_error_percent(predicted: float, reference: float) -> float:
    return 100.0 * (predicted - reference) / reference


def dataframe_to_markdown(frame: pd.DataFrame) -> str:
    """Small dependency-free Markdown table writer."""
    columns = [str(column) for column in frame.columns]
    rows = ["| " + " | ".join(columns) + " |", "| " + " | ".join(["---"] * len(columns)) + " |"]
    for values in frame.itertuples(index=False, name=None):
        formatted = []
        for value in values:
            if isinstance(value, (float, np.floating)):
                formatted.append(f"{float(value):.8g}")
            else:
                formatted.append(str(value))
        rows.append("| " + " | ".join(formatted) + " |")
    return "\n".join(rows)


def run_nominal_models() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    solutions = solve_models(
        WAVELENGTHS_NM,
        gap1_um=GAP1_UM,
        gap2_um=GAP2_UM,
        core_radius_um=CORE_RADIUS_UM,
        wall_thickness_um=WALL_THICKNESS_UM,
    )
    solution_df = pd.DataFrame([solution.to_dict() for solution in solutions])
    solution_df["layers"] = solution_df["model_name"].str.extract(r"^(\d+-layer)")

    bache_rows: list[dict] = []
    for wavelength_nm in WAVELENGTHS_NM:
        n_silica = silica_index_sellmeier(wavelength_nm * 1e-9)
        stacks = {
            "single wall": ([n_silica], [WALL_THICKNESS_UM]),
            "double wall radial stack": (
                [n_silica, 1.0, n_silica],
                [WALL_THICKNESS_UM, GAP1_UM, WALL_THICKNESS_UM],
            ),
            "triple wall radial stack": (
                [n_silica, 1.0, n_silica, 1.0, n_silica],
                [
                    WALL_THICKNESS_UM,
                    GAP1_UM,
                    WALL_THICKNESS_UM,
                    GAP2_UM,
                    WALL_THICKNESS_UM,
                ],
            ),
        }
        for stack_name, (indices, thicknesses) in stacks.items():
            for polarization in ("TE", "TM", "hybrid"):
                bache_rows.append(
                    {
                        "wavelength_nm": wavelength_nm,
                        "stack": stack_name,
                        "polarization": polarization,
                        "loss_db_km": multilayer_bouncing_ray_loss_db_km(
                            wavelength_nm,
                            indices,
                            thicknesses,
                            polarization=polarization,
                            core_radius_um=CORE_RADIUS_UM,
                        ),
                        "f_FEM": None,
                        "target_used_for_fitting": False,
                    }
                )
        bache_rows.append(
            {
                "wavelength_nm": wavelength_nm,
                "stack": "Bache Eq. 15-17 single wall",
                "polarization": "hybrid",
                "loss_db_km": bache_bouncing_ray_loss_db_km(
                    wavelength_nm,
                    polarization="hybrid",
                    core_radius_um=CORE_RADIUS_UM,
                    wall_thickness_um=WALL_THICKNESS_UM,
                ),
                "f_FEM": None,
                "target_used_for_fitting": False,
            }
        )
    bache_df = pd.DataFrame(bache_rows)

    comparison_rows: list[dict] = []
    for wavelength_nm in WAVELENGTHS_NM:
        reference = PETROVICH_MEASURED_TOTAL_DB_KM[wavelength_nm]
        for layer in ("5-layer", "7-layer"):
            predicted = float(
                solution_df.query(
                    "wavelength_nm == @wavelength_nm and layers == @layer"
                )["loss_db_km"].iloc[0]
            )
            error = signed_error_percent(predicted, reference)
            comparison_rows.append(
                {
                    "wavelength_nm": wavelength_nm,
                    "model_stage": "Raw" if layer == "5-layer" else "Current no-fit",
                    "model": (
                        "5-layer concentric scalar"
                        if layer == "5-layer"
                        else "7-layer concentric scalar (closest no-fit radial surrogate)"
                    ),
                    "predicted_leakage_db_km": predicted,
                    "paper_measured_total_db_km": reference,
                    "signed_error_percent": error,
                    "absolute_error_percent": abs(error),
                    "prediction_to_measurement_ratio": predicted / reference,
                    "within_10_percent": abs(error) <= 10.0,
                    "within_15_percent": abs(error) <= 15.0,
                    "target_used_for_fitting": TARGET_USED_FOR_FITTING,
                    "validation_split": "independent hold-out benchmark",
                    "calibrated_model_status": "not produced (Petrovich fitting prohibited)",
                    "quantity_warning": "leakage-only model compared with measured total loss",
                }
            )
    comparison_df = pd.DataFrame(comparison_rows)
    return solution_df, bache_df, comparison_df


def run_geometry_range_sweep() -> pd.DataFrame:
    """Propagate published diameter ranges without choosing a target-matching point."""
    values = lambda bounds: np.linspace(bounds[0], bounds[1], 3)
    rows: list[dict] = []
    for large in values(LARGE_DIAMETER_RANGE_UM):
        for middle in values(MIDDLE_DIAMETER_RANGE_UM):
            for small in values(SMALL_DIAMETER_RANGE_UM):
                gap1 = (large - 2 * WALL_THICKNESS_UM) - middle
                gap2 = (middle - 2 * WALL_THICKNESS_UM) - small
                if gap1 <= 0 or gap2 <= 0:
                    continue
                for wavelength_nm in WAVELENGTHS_NM:
                    model = build_models(
                        wavelength_nm,
                        gap1_um=float(gap1),
                        gap2_um=float(gap2),
                        core_radius_um=CORE_RADIUS_UM,
                        wall_thickness_um=WALL_THICKNESS_UM,
                    )["7-layer"]
                    solution = solve_fundamental_mode(wavelength_nm, model)
                    rows.append(
                        {
                            "wavelength_nm": wavelength_nm,
                            "large_diameter_um": large,
                            "middle_diameter_um": middle,
                            "small_diameter_um": small,
                            "gap1_um": gap1,
                            "gap2_um": gap2,
                            "converged": solution.converged,
                            "loss_db_km": solution.loss_db_km,
                            "characteristic_abs": solution.characteristic_abs,
                            "relative_smin": solution.relative_smallest_singular_value,
                            "root_message": solution.root_message,
                            "target_used_for_selection": False,
                        }
                    )
    return pd.DataFrame(rows)


def make_figures(
    solution_df: pd.DataFrame,
    comparison_df: pd.DataFrame,
    sweep_df: pd.DataFrame,
) -> None:
    plt.style.use("seaborn-v0_8-whitegrid")

    fig, ax = plt.subplots(figsize=(9.2, 5.1), layout="constrained")
    labels = ["1310 nm", "1550 nm"]
    x = np.arange(len(labels))
    width = 0.25
    paper = [PETROVICH_MEASURED_TOTAL_DB_KM[w] for w in WAVELENGTHS_NM]
    five = [
        float(comparison_df.query("wavelength_nm == @w and model.str.startswith('5-layer')", engine="python").predicted_leakage_db_km.iloc[0])
        for w in WAVELENGTHS_NM
    ]
    seven = [
        float(comparison_df.query("wavelength_nm == @w and model.str.startswith('7-layer')", engine="python").predicted_leakage_db_km.iloc[0])
        for w in WAVELENGTHS_NM
    ]
    ax.bar(x - width, paper, width, label="Petrovich HCF2 measured total")
    ax.bar(x, five, width, label="5-layer scalar leakage")
    ax.bar(x + width, seven, width, label="7-layer scalar leakage")
    ax.set_yscale("log")
    ax.set_xticks(x, labels)
    ax.set_ylabel("Loss (dB/km, log scale)")
    ax.set_title("No-fit code prediction versus Petrovich HCF2 measurement")
    ax.legend(fontsize=9)
    fig.savefig(OUTPUT_DIR / "01_no_fit_loss_comparison.png", dpi=180)
    fig.savefig(OUTPUT_DIR / "01_no_fit_loss_comparison.svg")
    plt.close(fig)

    closest = comparison_df[comparison_df.model.str.startswith("7-layer")]
    fig, ax = plt.subplots(figsize=(8.2, 4.8), layout="constrained")
    bars = ax.bar(labels, closest.absolute_error_percent, color=["#d95f02", "#7570b3"])
    ax.axhline(10, color="red", linestyle="--", label="10% threshold")
    ax.axhline(15, color="darkorange", linestyle=":", label="15% threshold")
    ax.set_yscale("log")
    ax.set_ylabel("Absolute percentage error (%)")
    ax.set_title("Closest no-fit radial surrogate still misses the target")
    for bar, value in zip(bars, closest.absolute_error_percent):
        ax.text(bar.get_x() + bar.get_width() / 2, value * 1.07, f"{value:.1f}%", ha="center")
    ax.legend()
    fig.savefig(OUTPUT_DIR / "02_absolute_error_percent.png", dpi=180)
    fig.savefig(OUTPUT_DIR / "02_absolute_error_percent.svg")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(8.8, 5.0), layout="constrained")
    for layer, marker in (("3-layer", "o"), ("5-layer", "s"), ("7-layer", "^")):
        part = solution_df[solution_df.layers == layer]
        ax.plot(part.wavelength_nm, part.loss_db_km, marker=marker, linewidth=2, label=layer)
    ax.set_yscale("log")
    ax.set_xlabel("Wavelength (nm)")
    ax.set_ylabel("Scalar leakage loss (dB/km)")
    ax.set_title("Radial nesting reduces leakage, but omits azimuthal confinement")
    ax.legend()
    fig.savefig(OUTPUT_DIR / "03_radial_nesting_reduction.png", dpi=180)
    fig.savefig(OUTPUT_DIR / "03_radial_nesting_reduction.svg")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(8.8, 5.0), layout="constrained")
    data = [
        sweep_df.query("wavelength_nm == @w and converged").loss_db_km.to_numpy()
        for w in WAVELENGTHS_NM
    ]
    ax.boxplot(data, tick_labels=labels, showfliers=True)
    ax.scatter([1, 2], paper, marker="x", s=90, linewidth=2.5, color="red", label="paper total loss")
    ax.set_yscale("log")
    ax.set_ylabel("7-layer leakage loss (dB/km)")
    ax.set_title("Published diameter-range propagation (27 fixed grid corners/midpoints)")
    ax.legend()
    fig.savefig(OUTPUT_DIR / "04_geometry_range_sensitivity.png", dpi=180)
    fig.savefig(OUTPUT_DIR / "04_geometry_range_sensitivity.svg")
    plt.close(fig)


def main() -> None:
    started = time.perf_counter()
    solution_df, bache_df, comparison_df = run_nominal_models()
    sweep_df = run_geometry_range_sweep()
    runtime_seconds = time.perf_counter() - started

    assert solution_df.converged.all(), "a nominal complex root did not converge"
    assert not TARGET_USED_FOR_FITTING

    # Independent mathematical validation of the single-wall scalar solver.
    validation_rows = []
    for wavelength_nm in WAVELENGTHS_NM:
        tmm = float(
            solution_df.query(
                "wavelength_nm == @wavelength_nm and layers == '3-layer'"
            ).loss_db_km.iloc[0]
        )
        closed = bouncing_ray_te_loss_db_km(
            wavelength_nm,
            core_radius_um=CORE_RADIUS_UM,
            wall_thickness_um=WALL_THICKNESS_UM,
        )
        error = signed_error_percent(tmm, closed)
        validation_rows.append(
            {
                "wavelength_nm": wavelength_nm,
                "three_layer_tmm_db_km": tmm,
                "bache_te_closed_form_db_km": closed,
                "signed_error_percent": error,
                "absolute_error_percent": abs(error),
            }
        )
    validation_df = pd.DataFrame(validation_rows)
    assert validation_df.absolute_error_percent.max() < 10.0

    sweep_summary_df = (
        sweep_df[sweep_df.converged].groupby("wavelength_nm").loss_db_km
        .agg(["min", "median", "max"])
        .reset_index()
    )
    sweep_summary_df["paper_measured_total_db_km"] = sweep_summary_df.wavelength_nm.map(
        PETROVICH_MEASURED_TOTAL_DB_KM
    )

    make_figures(solution_df, comparison_df, sweep_df)

    solution_df.to_csv(OUTPUT_DIR / "01_nominal_mode_solutions.csv", index=False)
    validation_df.to_csv(OUTPUT_DIR / "02_three_layer_math_validation.csv", index=False)
    bache_df.to_csv(OUTPUT_DIR / "03_bache_multilayer_no_fit.csv", index=False)
    comparison_df.to_csv(OUTPUT_DIR / "04_petrovich_no_fit_comparison.csv", index=False)
    sweep_df.to_csv(OUTPUT_DIR / "05_geometry_range_sweep.csv", index=False)
    sweep_summary_df.to_csv(OUTPUT_DIR / "06_geometry_range_summary.csv", index=False)

    five_mape = float(
        comparison_df[comparison_df.model.str.startswith("5-layer")].absolute_error_percent.mean()
    )
    seven_mape = float(
        comparison_df[comparison_df.model.str.startswith("7-layer")].absolute_error_percent.mean()
    )
    summary = {
        "runtime_seconds": runtime_seconds,
        "python": sys.version.split()[0],
        "numpy": np.__version__,
        "scipy": scipy.__version__,
        "platform": platform.platform(),
        "target_used_for_fitting": TARGET_USED_FOR_FITTING,
        "nominal_core_radius_um": CORE_RADIUS_UM,
        "nominal_wall_thickness_um": WALL_THICKNESS_UM,
        "derived_gap1_um": GAP1_UM,
        "derived_gap2_um": GAP2_UM,
        "three_layer_math_validation_mape_percent": float(
            validation_df.absolute_error_percent.mean()
        ),
        "five_layer_petrovich_mape_percent": five_mape,
        "seven_layer_petrovich_mape_percent": seven_mape,
        "ten_percent_target_passed": bool(
            comparison_df[comparison_df.model.str.startswith("7-layer")]
            .absolute_error_percent.le(10.0)
            .all()
        ),
        "fifteen_percent_target_passed": bool(
            comparison_df[comparison_df.model.str.startswith("7-layer")]
            .absolute_error_percent.le(15.0)
            .all()
        ),
        "all_roots_converged": bool(solution_df.converged.all() and sweep_df.converged.all()),
        "geometry_sweep_points": int(len(sweep_df)),
        "geometry_sweep_nonconverged": int((~sweep_df.converged).sum()),
    }
    (OUTPUT_DIR / "summary.json").write_text(
        json.dumps(summary, indent=2), encoding="utf-8"
    )

    closest = comparison_df[comparison_df.model.str.startswith("7-layer")]
    report_lines = [
        "# 20260908 HCF2 no-fit loss validation",
        "",
        "Petrovich HCF2 loss values were used only after prediction as benchmark values.",
        "No Petrovich point, fitted multiplier, regression coefficient, or target-selected geometry entered the solver.",
        "",
        "## Executed result",
        "",
        dataframe_to_markdown(closest),
        "",
        f"- 3-layer TMM vs Bache TE mathematical-validation MAPE: {summary['three_layer_math_validation_mape_percent']:.3f}%",
        f"- 5-layer vs Petrovich MAPE: {five_mape:.3f}%",
        f"- 7-layer vs Petrovich MAPE: {seven_mape:.3f}%",
        f"- 10% target passed: {summary['ten_percent_target_passed']}",
        f"- 15% target passed: {summary['fifteen_percent_target_passed']}",
        "",
        "## Required improvement when error exceeds 10%",
        "",
        "1. Replace the one-dimensional concentric surrogate with a full-vector 2D complex-eigenvalue solver on the actual five-tube, double-nested SEM contour; use a converged mesh and PML study.",
        "2. Include azimuthal confinement and radial glass webs; a radial TMM cannot represent the field shaping identified by Murphy and Bird (2023).",
        "3. Predict total loss as LL + SSL + microbend + gas absorption. The present code predicts leakage only, whereas the reference is measured total attenuation.",
        "4. Acquire membrane-thickness, inter-tube-gap, ellipticity, angular offset, longitudinal-variation, surface-roughness PSD, coating/outer-diameter and spool/bend inputs rather than midpoint radial gaps.",
        "5. Validate on fibres/wavelengths excluded from calibration. If empirical SSL/microbend coefficients are used, label the result calibrated rather than geometry-only.",
    ]
    (OUTPUT_DIR / "20260908_results.md").write_text(
        "\n".join(report_lines), encoding="utf-8"
    )

    print(json.dumps(summary, indent=2))
    print("\nClosest no-fit comparison:\n", closest.to_string(index=False))
    print("\nGeometry-range summary:\n", sweep_summary_df.to_string(index=False))


if __name__ == "__main__":
    main()
