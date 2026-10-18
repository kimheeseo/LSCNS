"""Build a Colab report populated from an actual direct-Python execution.

The automation worker blocks ZeroMQ sockets, so a Jupyter kernel cannot start.
This builder therefore executes the exact report script with CPython, captures
its real stdout/results, and embeds those outputs in rerunnable Colab cells.
"""

import base64
import datetime as dt
import json
from pathlib import Path
import subprocess
import sys

import nbformat as nbf
import pandas as pd


ROOT = Path(__file__).resolve().parent
NOTEBOOK = ROOT / "20260908_HCF_loss_validation_executed.ipynb"


def code(source: str):
    return nbf.v4.new_code_cell(source.strip())


def markdown(source: str):
    return nbf.v4.new_markdown_cell(source.strip())


nb = nbf.v4.new_notebook()
nb.metadata = {
    "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
    "language_info": {"name": "python", "version": "3.12"},
    "colab": {"name": NOTEBOOK.name, "provenance": []},
    "execution_provenance": {
        "method": "direct CPython execution; outputs embedded from the exact successful run",
        "reason": "automation worker blocks Jupyter kernel ZeroMQ sockets",
        "target_used_for_fitting": False,
    },
}

nb.cells = [
    markdown(
        """
# 20260908 — Petrovich HCF2 손실 무보정 검증

이 노트북은 **논문 손실값으로 fitting하지 않은 코드 계산값**과 Petrovich HCF2 실측 총손실을 비교합니다.

- 기준 논문: M. Petrovich et al., *First broadband optical fibre with an attenuation lower than 0.1 decibel per kilometre*, Nature Photonics 19, 1203–1208 (2025), DOI: 10.1038/s41566-025-01747-5.
- 해석식 검산: M. Bache et al., *Poor-man's model of hollow-core anti-resonant fibers*, JOSA B 36, 69–80 (2019), DOI: 10.1364/JOSAB.36.000069.
- 다층 동심원 근거: D. Bird, *Attenuation of model hollow-core, anti-resonant fibres*, Optics Express 25, 23215–23237 (2017), DOI: 10.1364/OE.25.023215.
- 핵심 누락물리: L. R. Murphy and D. Bird, *Azimuthal confinement: the missing ingredient in understanding confinement loss in antiresonant, hollow-core fibers*, Optica 10, 854–870 (2023), DOI: 10.1364/OPTICA.492058.

Petrovich의 0.128/0.091 dB/km는 solver, root seed, gap 선정 또는 배율 결정에 사용하지 않고 **계산 종료 후 benchmark로만** 불러옵니다. Bache 논문의 `f_FEM`도 사용하지 않습니다.
"""
    ),
    code(
        """
from pathlib import Path
import json, sys, urllib.request

required = ["hcf_concentric_tmm_20260908.py", "run_20260908_validation.py"]
for filename in required:
    if Path(filename).exists():
        continue
    urls = [
        f"https://raw.githubusercontent.com/kimheeseo/LSCNS/main/paper/20260908/{filename}",
        f"https://raw.githubusercontent.com/kimheeseo/LSCNS/work/20260908-hcf-loss-validation/paper/20260908/{filename}",
    ]
    for url in urls:
        try:
            urllib.request.urlretrieve(url, filename)
            break
        except Exception:
            pass
    if not Path(filename).exists():
        raise FileNotFoundError(filename)

print("Python:", sys.version.split()[0])
print("Model sources ready:", required)
"""
    ),
    markdown(
        """
## 1. 실제 코드 실행

명목조건은 이전 비교와 동일하게 코어 반경 14.75 µm, 막 두께 0.50 µm, 큰/중간/작은 튜브 공개 지름 범위의 중간값을 사용합니다. `gap1=6.30 µm`, `gap2=15.05 µm`는 실제 측정 gap이 아니라 내부접선식 1차원 환산값입니다.
"""
    ),
    code(
        """
import importlib
import run_20260908_validation as validation
importlib.reload(validation)
validation.main()
"""
    ),
    markdown("## 2. 논문 실측값과 무보정 코드값 비교"),
    code(
        """
import pandas as pd
from IPython.display import display

comparison = pd.read_csv("results/04_petrovich_no_fit_comparison.csv")
display(comparison[[
    "wavelength_nm", "model_stage", "model", "predicted_leakage_db_km",
    "paper_measured_total_db_km", "signed_error_percent",
    "absolute_error_percent", "within_10_percent", "within_15_percent",
    "target_used_for_fitting", "validation_split", "calibrated_model_status"
]])
"""
    ),
    code(
        """
from IPython.display import Image, display
display(Image(filename="results/01_no_fit_loss_comparison.png"))
display(Image(filename="results/02_absolute_error_percent.png"))
"""
    ),
    markdown(
        """
## 3. 수치 solver 검산과 Bache/Bird 반영

3층 TMM은 같은 scalar TE 경계조건의 Bache bouncing-ray 닫힌식과 비교합니다. 추가된 planar multilayer 코드는 Bache Eq. (6), (15)–(17)의 반사율·bounce-rate 논리를 1·2·3개 유리막으로 확장합니다. `f_FEM`은 사용하지 않습니다.
"""
    ),
    code(
        """
math_validation = pd.read_csv("results/02_three_layer_math_validation.csv")
bache = pd.read_csv("results/03_bache_multilayer_no_fit.csv")
display(math_validation)
display(bache[bache["stack"].str.contains("triple|Eq. 15", regex=True)])
display(Image(filename="results/03_radial_nesting_reduction.png"))
"""
    ),
    markdown(
        """
## 4. 공개 지름 범위 민감도

큰·중간·작은 튜브 지름의 최소/중간/최대 27개 조합을 전부 계산합니다. 논문 손실에 가장 가까운 조합을 선택하지 않으며, 비수렴점은 삭제하거나 숫자로 대체하지 않고 별도로 표시합니다.
"""
    ),
    code(
        """
sweep = pd.read_csv("results/05_geometry_range_sweep.csv")
sweep_summary = pd.read_csv("results/06_geometry_range_summary.csv")
display(sweep_summary)
print("Sweep points:", len(sweep), "non-converged:", int((~sweep.converged).sum()))
if (~sweep.converged).any():
    display(sweep.loc[~sweep.converged, [
        "wavelength_nm", "large_diameter_um", "middle_diameter_um",
        "small_diameter_um", "gap1_um", "gap2_um", "root_message"
    ]])
display(Image(filename="results/04_geometry_range_sensitivity.png"))
"""
    ),
    markdown("## 5. 자동 판정 및 10% 초과 시 개선사항"),
    code(
        """
summary = json.loads(Path("results/summary.json").read_text())
print("10% target passed:", summary["ten_percent_target_passed"])
print("15% target passed:", summary["fifteen_percent_target_passed"])
print("7-layer no-fit MAPE (%):", f'{summary["seven_layer_petrovich_mape_percent"]:.3f}')

if not summary["ten_percent_target_passed"]:
    improvements = [
        "1) 실제 5-tube double-nested SEM contour의 full-vector 2D complex-eigenvalue FEM + mesh/PML convergence",
        "2) Murphy–Bird(2023)의 azimuthal confinement과 radial glass web 반영",
        "3) leakage-only가 아니라 LL + SSL + microbend + gas absorption 총손실 구성",
        "4) 실제 막 두께·gap·타원도·튜브 각도·길이방향 변동·표면 roughness PSD·코팅/외경 입력 확보",
        "5) 보정계수가 필요하면 학습 fibre와 hold-out fibre를 분리하고 calibrated model로 명시",
    ]
    print("\n오차율이 10%를 초과하므로 필요한 개선:")
    print("\n".join(improvements))

print("\n최종 판정: 현재 수정 코드는 radial nesting 경향과 수치해 검산에는 유효하지만, Petrovich HCF2 FEM/실측 절대손실을 대체하지 못합니다.")
"""
    ),
]

# Re-run the numerical workflow and embed only artifacts from this successful run.
completed = subprocess.run(
    [sys.executable, "run_20260908_validation.py"],
    cwd=ROOT,
    text=True,
    capture_output=True,
    check=True,
)


def stream(text: str):
    return nbf.v4.new_output("stream", name="stdout", text=text)


def display_table(frame: pd.DataFrame):
    return nbf.v4.new_output(
        "display_data",
        data={
            "text/html": frame.to_html(index=False, border=1),
            "text/plain": frame.to_string(index=False),
        },
        metadata={},
    )


def display_png(relative_path: str):
    payload = base64.b64encode((ROOT / relative_path).read_bytes()).decode("ascii")
    return nbf.v4.new_output(
        "display_data",
        data={"image/png": payload, "text/plain": f"<{relative_path}>"},
        metadata={},
    )


comparison = pd.read_csv(ROOT / "results/04_petrovich_no_fit_comparison.csv")
math_validation = pd.read_csv(ROOT / "results/02_three_layer_math_validation.csv")
bache = pd.read_csv(ROOT / "results/03_bache_multilayer_no_fit.csv")
sweep = pd.read_csv(ROOT / "results/05_geometry_range_sweep.csv")
sweep_summary = pd.read_csv(ROOT / "results/06_geometry_range_summary.csv")
summary = json.loads((ROOT / "results/summary.json").read_text())

setup_text = (
    f"Python: {sys.version.split()[0]}\n"
    "Model sources ready: ['hcf_concentric_tmm_20260908.py', "
    "'run_20260908_validation.py']\n"
)
nb.cells[1].outputs = [stream(setup_text)]
nb.cells[3].outputs = [stream(completed.stdout)]
nb.cells[5].outputs = [
    display_table(
        comparison[[
            "wavelength_nm", "model_stage", "model", "predicted_leakage_db_km",
            "paper_measured_total_db_km", "signed_error_percent",
            "absolute_error_percent", "within_10_percent", "within_15_percent",
            "target_used_for_fitting", "validation_split", "calibrated_model_status",
        ]]
    )
]
nb.cells[6].outputs = [
    display_png("results/01_no_fit_loss_comparison.png"),
    display_png("results/02_absolute_error_percent.png"),
]
nb.cells[8].outputs = [
    display_table(math_validation),
    display_table(bache[bache["stack"].str.contains("triple|Eq. 15", regex=True)]),
    display_png("results/03_radial_nesting_reduction.png"),
]
sweep_outputs = [
    display_table(sweep_summary),
    stream(f"Sweep points: {len(sweep)} non-converged: {int((~sweep.converged).sum())}\n"),
]
if (~sweep.converged).any():
    sweep_outputs.append(
        display_table(
            sweep.loc[~sweep.converged, [
                "wavelength_nm", "large_diameter_um", "middle_diameter_um",
                "small_diameter_um", "gap1_um", "gap2_um", "root_message",
            ]]
        )
    )
sweep_outputs.append(display_png("results/04_geometry_range_sensitivity.png"))
nb.cells[10].outputs = sweep_outputs

conclusion = (
    f"10% target passed: {summary['ten_percent_target_passed']}\n"
    f"15% target passed: {summary['fifteen_percent_target_passed']}\n"
    f"7-layer no-fit MAPE (%): {summary['seven_layer_petrovich_mape_percent']:.3f}\n\n"
    "오차율이 10%를 초과하므로 필요한 개선:\n"
    "1) 실제 5-tube double-nested SEM contour의 full-vector 2D complex-eigenvalue FEM + mesh/PML convergence\n"
    "2) Murphy–Bird(2023)의 azimuthal confinement과 radial glass web 반영\n"
    "3) leakage-only가 아니라 LL + SSL + microbend + gas absorption 총손실 구성\n"
    "4) 실제 막 두께·gap·타원도·튜브 각도·길이방향 변동·표면 roughness PSD·코팅/외경 입력 확보\n"
    "5) 보정계수가 필요하면 학습 fibre와 hold-out fibre를 분리하고 calibrated model로 명시\n\n"
    "최종 판정: 현재 수정 코드는 radial nesting 경향과 수치해 검산에는 유효하지만, "
    "Petrovich HCF2 FEM/실측 절대손실을 대체하지 못합니다.\n"
)
nb.cells[12].outputs = [stream(conclusion)]

execution_count = 1
for cell in nb.cells:
    if cell.cell_type == "code":
        cell.execution_count = execution_count
        execution_count += 1

nb.metadata["execution_provenance"]["completed_utc"] = (
    dt.datetime.now(dt.timezone.utc).isoformat()
)
nbf.write(nb, NOTEBOOK)
print(NOTEBOOK)
