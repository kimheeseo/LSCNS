## 2026-11-03 추가: LE125 굽힘손실 2D FEM/PML 계산·사양 비교

**[R=30 mm · 100 turns · 1550/1625 nm 상세 비교 보고서](./LE125_BENDING_VALIDATION_2026-11-03.md)**

- [동일 코드 기반 Colab 재실행 노트북](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/validation_report/LS_PremierLinkLE125_G.654.E/LE125_Bending_Validation_2026_11_03.ipynb)
- [1550 nm 2D FEM/PML 시연 JSON](./LE125_BEND_1550_UNVALIDATED.json) · [1625 nm 2D FEM/PML 시연 JSON](./LE125_BEND_1625_UNVALIDATED.json)
- [GitHub Pages의 2D FEM 결과 가져오기](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/?v=20261103-complex-pml-1#bend2d-fem)

**제공된 LE125 사양표:** 1550 / 1625 nm 각각 반경 **30 mm × 100 turns**에서 추가 굽힘손실 **≤0.1 dB**.

**가상 Step-index 계산값:** 1550 nm **0.9741 dB/100 turns**, 1625 nm **4.9800 dB/100 turns** (h≈0.55 µm, PML=7λ, 계산창 63 µm). **두 파장 모두 PML 수렴성 검증 실패**, 1550 nm 메시 수렴성도 미달입니다. 실제 LE125 굴절률 분포(RIP), 코팅/외부 경계, 편광·GEBT 응력장이 미확보·미구현이므로 **실제 제품 합격/불합격은 판정할 수 없습니다**. 숫자와 사양의 차이는 **근사 모델 불일치**로 해석해야 합니다.

---

# G.654.E / PremierLink LE125: numerical benchmark (2026-10-19)

## Google Colab에서 직접 시뮬레이션하기

[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/validation_report/LS_PremierLinkLE125_G.654.E/G654E_Colab_Validation.ipynb)

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
