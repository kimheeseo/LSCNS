# G.654.E / PremierLink LE125: numerical benchmark (2026-10-19)

## Google Colab에서 직접 시뮬레이션하기

[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/g654e_validation/G654E_Colab_Validation.ipynb)

- [실행 결과(그래프 포함)를 저장한 노트북](G654E_Colab_Validation.ipynb): Colab의 `런타임 → 모두 실행` 또는 각 셀의 실행 버튼으로 실제 1D radial FEM, 2D P1 FEM과 분산 그래프를 계산합니다.
- [기본 실행 결과 ZIP](G654E_Colab_Results.zip): 저장된 그래프 8개, CSV 표, 실행 보고서, JSON. 파라미터를 바꿨다면 Colab에서 다시 계산 후 새 ZIP을 생성하세요.
- **필수 주의:** 공개 12.5 µm MFD는 *보정 목표*이며 실제 제품의 코어 구조, Aeff, 감쇠를 실측값과 독립 비교한 검증이 아닙니다. 처음 입력된 가상 step-index 모델에서 11/12개 수치 검사가 통과합니다.


This folder contains the reproducible G.654.E **step-index surrogate** numerical consistency benchmark for the existing M7 scalar fiber FEM code. It does **not** certify the LS Cable & System PremierLink LE125 product or COMSOL equivalence.

## Summary
- Calibrated to published **12.5 µm MFD at 1550 nm** (MFD is an input fit target, not an independent validation).
- Assumed 5.5 µm core radius; fitted index step Δn ≈ 0.003478579.
- Bessel baseline: neff ≈ 1.445718991, Aeff ≈ 121.923 µm², D ≈ 19.826 ps/(nm·km), dispersion slope S ≈ 0.05934 ps/(nm²·km).
- **11/12** predefined numerical checks pass. The original uniform-grid 2D FEM misses a strict neff tolerance at the 161×161 grid and shows non-monotonic mesh convergence.
- Full-vector Maxwell, true manufacturer refractive-index profile, attenuation, cutoff certification, and PML leakage are **not** verified.

## Files
- [Validation report](G654E_VALIDATION_REPORT.md) — methodology, limits, tables, sources and figures.
- [Radial FEM and Bessel benchmark](g654e_validation.py) — calibrated index step, spectral dispersion, radial mesh study.
- [Original 2D FEM method cross-check](validate_m7_2d.py) — triangle-centroid material tags, mesh-error evaluation.
- [JSON results](g654e_results.json), [CSV summary](g654e_summary.csv).
- [Dispersion plot](g654e_dispersion.png), [Mode and radial convergence](g654e_mode_convergence.png), [2D mesh error](g654e_2d_mesh_error.png).
- [Full reproducibility ZIP](G654E_M7_validation_package.zip).

## Reproduce
Requires Python 3.11+:

```bash
python -m pip install numpy scipy matplotlib
python g654e_validation.py
python validate_m7_2d.py
```

**Do not report fitted MFD as an independently predicted G.654.E product property.** The source scripts transpose relevant M7 numerical routines for independent benchmark execution; they do not implement a commercial COMSOL solver.
