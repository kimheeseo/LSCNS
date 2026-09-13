# EGN_model — EGN Validation and GN/EGN Comparison

이 폴더는 EGN(Enhanced Gaussian Noise) 모델 구현과 검증 자료를 두 개의 하위 폴더로 구분하여 정리한 공간입니다. `EGN_Model_Validation`은 Carena 2014 논문의 SCI/XCI/MCI 결과를 직접 검증하는 데 초점을 두며, `EGNvsGN`은 범용 GN 적분 엔진과 변조 성능 레이어를 논문의 GN/EGN/SIM 곡선과 비교하는 데 초점을 둡니다.

**English:** This directory organizes EGN (Enhanced Gaussian Noise) implementation and validation work into two subdirectories. `EGN_Model_Validation` focuses on direct validation of SCI/XCI/MCI behavior against Carena 2014, while `EGNvsGN` focuses on comparing the general GN-integral engine and modulation-performance layer with published GN/EGN/SIM curves.

---

## 1. 전체 구조 / Directory structure

```text
EGN_model/
├─ EGN_Model_Validation/
│  ├─ final_EGN.py
│  ├─ final_EGN_Carena_Fig1_3_6_8_Validation_Colab.ipynb
│  ├─ README_final_EGN_validation.md
│  ├─ final_EGN_validation_summary.csv
│  ├─ final_EGN_paper_vs_code_detail.csv
│  ├─ Fig1_SMF_paper_vs_code.png
│  ├─ Fig3_SMF_paper_vs_code.png
│  ├─ Fig6_SMF_paper_vs_code.png
│  ├─ Fig8_SMF_paper_vs_code.png
│  ├─ quadrature_sensitivity_SMF_50span.png
│  └─ final_EGN_validation_bundle.zip
│
└─ EGNvsGN/
   ├─ GN_integral_math_verification.ipynb
   ├─ GN_integral_usage_guide.ipynb
   ├─ GN_modulation_Fig123_validation_colab.ipynb
   ├─ paper_fig12_reference.json
   ├─ paper_fig3_reference.json
   ├─ fig12_results_avg.json
   ├─ fig3_code_gn_seeds.json
   └─ readme.md
```

두 폴더는 비슷해 보이지만 목적이 다릅니다. `EGN_Model_Validation`은 full-EGN 알고리즘 자체의 정확성과 수치 수렴성을 검증하고, `EGNvsGN`은 현재 범용 GN 엔진이 논문의 GN 결과를 얼마나 잘 재현하는지와 modulation/system layer가 어떤 범위까지 유효한지를 검증합니다.

**English:** Although the two subdirectories are related, their purposes differ. `EGN_Model_Validation` evaluates the accuracy and numerical convergence of the full-EGN implementation itself, whereas `EGNvsGN` evaluates how accurately the general GN engine reproduces published GN results and clarifies the valid scope of the modulation/system-performance layer.

---

# 2. `EGN_Model_Validation`

## 폴더 목적 / Purpose

A. Carena et al., **“EGN model of non-linear fiber propagation”** (Optics Express, 2014)의 Fig. 1, 3, 6, 8 조건을 이용하여 `final_EGN.py`의 GN baseline과 SCI/XCI/MCI EGN correction을 검증하기 위한 폴더입니다.

**English:** This folder validates the GN baseline and SCI/XCI/MCI EGN corrections implemented in `final_EGN.py` using the conditions of Figs. 1, 3, 6, and 8 from A. Carena et al., **“EGN model of non-linear fiber propagation”** (Optics Express, 2014).

검증 대상은 다음과 같습니다.

**English:** The validation targets are:

- Fig. 1: SCI  
  **English:** Fig. 1: SCI
- Fig. 3: 3-channel XCI only  
  **English:** Fig. 3: XCI-only, 3-channel case
- Fig. 6: 3-channel XMCI = XCI + MCI  
  **English:** Fig. 6: XMCI = XCI + MCI, 3-channel case
- Fig. 8: 9-channel XMCI = XCI + MCI  
  **English:** Fig. 8: XMCI = XCI + MCI, 9-channel case

### `final_EGN.py`

Poggiolini 계열 GN 적분을 baseline으로 사용하고, Carena 2014의 비가우시안 보정항을 SCI, XCI, MCI에 추가하도록 구성한 실험적 full-EGN 구현입니다. 코드에서는 GN baseline, SCI-EGN, full-EGN을 구분하여 사용할 수 있도록 설계되어 있습니다.

**English:** This is an experimental full-EGN implementation that uses a Poggiolini-style GN integral as the baseline and adds the non-Gaussian SCI, XCI, and MCI correction structure from Carena 2014. The implementation separates GN baseline, SCI-EGN, and full-EGN operating modes.

현재 상태에서 GN baseline과 SCI-EGN은 논문값과 좋은 일치를 보이지만, XCI/MCI correction integral은 quadrature order에 따른 수치 변화가 커서 완전히 수렴한 production-grade 모델로 간주하면 안 됩니다.

**English:** At the current stage, the GN baseline and SCI-EGN show good agreement with the paper, but the XCI/MCI correction integrals remain sensitive to quadrature order. Therefore, the full-EGN path should not yet be considered production-grade or fully converged.

### `final_EGN_Carena_Fig1_3_6_8_Validation_Colab.ipynb`

`final_EGN.py`를 Carena 2014 Fig. 1, 3, 6, 8 조건에서 실행하고, 논문에서 추출한 기준값과 코드 계산값을 그래프와 오차 지표로 비교하기 위한 Colab 검증 노트북입니다.

**English:** This Colab notebook runs `final_EGN.py` under the Carena 2014 Fig. 1, 3, 6, and 8 conditions and compares code results with paper-derived reference data using plots and error metrics.

주요 비교 지표는 dB 절대오차, MAE, 최대오차, 선형 `eta` 기준 상대오차입니다.

**English:** Main comparison metrics include absolute error in dB, mean absolute error (MAE), maximum error, and relative error based on linear `eta`.

### `README_final_EGN_validation.md`

전체 검증 조건, 결과, 현재 구현의 성숙도와 제한사항을 정리한 상세 validation report입니다. full-EGN이 완전히 검증된 것으로 오해하지 않도록 GN, SCI-EGN, XCI/MCI-EGN의 상태를 구분해 설명합니다.

**English:** This detailed validation report documents the test conditions, results, implementation maturity, and limitations. It explicitly separates the status of GN, SCI-EGN, and XCI/MCI-EGN to avoid presenting the full-EGN implementation as fully validated.

### `final_EGN_validation_summary.csv`

Fig. 1/3/6/8 및 fiber별 검증 결과를 요약한 표 형식 데이터입니다. 각 조건의 paper/code 비교와 대표 오차 지표를 빠르게 확인하기 위한 파일입니다.

**English:** This CSV summarizes validation results by figure and fiber, providing a compact paper-versus-code comparison and representative error metrics.

### `final_EGN_paper_vs_code_detail.csv`

각 span point 및 조건별 paper value, code value, 오차를 보다 세부적으로 기록한 데이터입니다. 평균값뿐 아니라 개별 point의 편차를 확인할 때 사용합니다.

**English:** This file stores detailed point-by-point paper values, code values, and errors for each span and test condition, allowing inspection beyond aggregate metrics.

### `Fig1_SMF_paper_vs_code.png`

SMF 조건의 Fig. 1 SCI에 대해 paper와 code를 시각적으로 비교한 그래프입니다.

**English:** Visual paper-versus-code comparison for Fig. 1 SCI under the SMF condition.

### `Fig3_SMF_paper_vs_code.png`

SMF 조건의 Fig. 3 XCI-only 결과를 비교하는 그래프입니다.

**English:** Visual comparison of the Fig. 3 XCI-only result for SMF.

### `Fig6_SMF_paper_vs_code.png`

SMF 조건의 Fig. 6 XMCI(XCI+MCI) 결과를 비교하는 그래프입니다.

**English:** Visual comparison of the Fig. 6 XMCI (XCI+MCI) result for SMF.

### `Fig8_SMF_paper_vs_code.png`

SMF 조건의 Fig. 8 9-channel XMCI 결과를 비교하는 그래프입니다.

**English:** Visual comparison of the Fig. 8 nine-channel XMCI result for SMF.

### `quadrature_sensitivity_SMF_50span.png`

XCI/MCI EGN correction에서 numerical quadrature order를 바꾸었을 때 결과가 어떻게 변하는지 보여주는 민감도 그래프입니다. 현재 full-EGN 구현의 가장 중요한 제한사항인 수치 수렴 문제를 확인하기 위한 자료입니다.

**English:** This sensitivity plot shows how the XCI/MCI EGN result changes with numerical quadrature order. It is specifically intended to expose the current full-EGN implementation's main limitation: numerical convergence of the correction integrals.

### `final_EGN_validation_bundle.zip`

검증 실행에 사용된 코드·데이터·결과물을 하나로 묶어 보관하기 위한 bundle입니다.

**English:** Archive bundle containing the validation code, data, and generated results for convenient preservation or transfer.

---

## 3. `EGN_Model_Validation` 성능 해석 / Validation status

저장된 validation report 기준으로 GN baseline은 Carena 논문의 Fig. 1/3/6/8 GN 값과 매우 잘 일치합니다. 대표 MAE는 약 **0.05–0.08 dB** 수준이며, SCI-EGN의 Fig. 1 전체 MAE는 약 **0.144 dB**로 보고되었습니다.

**English:** According to the stored validation report, the GN baseline agrees closely with the GN values in Carena Figs. 1/3/6/8, with representative MAEs of approximately **0.05–0.08 dB**. The reported overall Fig. 1 SCI-EGN MAE is approximately **0.144 dB**.

반면 Fig. 3/6/8의 full-EGN XCI/MCI는 특정 quadrature order에서 큰 변동을 보이며 비단조적인 수렴이 관찰됩니다. 따라서 현재 권장 상태는 다음과 같습니다.

**English:** By contrast, the full-EGN XCI/MCI paths for Figs. 3/6/8 exhibit significant quadrature-order sensitivity and non-monotonic convergence. The recommended interpretation is therefore:

```text
GN baseline : VALIDATED
SCI-EGN     : VALIDATED / RESEARCH-GRADE
XCI-EGN     : IMPLEMENTED, convergence improvement required
MCI-EGN     : IMPLEMENTED, convergence improvement required
Full-EGN    : NOT YET PRODUCTION-GRADE
```

이는 XCI/MCI 수식 구조가 구현되지 않았다는 의미가 아니라, 현재 numerical integration이 설계 보증에 사용할 정도로 충분히 안정적으로 수렴하지 않았다는 의미입니다.

**English:** This does not mean the XCI/MCI formula structure is absent; it means that the current numerical integration is not yet sufficiently stable and converged for design sign-off or specification guarantees.

---

# 4. `EGNvsGN`

## 폴더 목적 / Purpose

이 폴더는 `gn_integral_general.py`와 `gn_integral_general_modulation.py`의 GN 적분 경로를 수학적·수치적으로 확인하고, **“A Simple and Accurate Closed-Form EGN Model Formula”**의 Fig. 1–3 조건에서 paper GN/EGN/SIM 결과와 비교하기 위한 검증 자료를 포함합니다.

**English:** This directory verifies the mathematical and numerical behavior of the GN-integral path implemented in `gn_integral_general.py` and `gn_integral_general_modulation.py`, and compares it with the paper GN/EGN/SIM results of Figs. 1–3 from **“A Simple and Accurate Closed-Form EGN Model Formula.”**

현재 코드가 직접 재현하는 핵심 대상은 paper GN 곡선입니다. 논문의 EGN과 SSFM/SIM 곡선은 같은 그래프에서 참고값으로 표시하지만, 현재 범용 엔진이 full-WDM EGN 또는 SSFM을 구현했다고 해석하면 안 됩니다.

**English:** The primary directly reproducible target is the paper GN curve. Published EGN and SSFM/SIM curves are retained as reference curves, but the general engine should not be interpreted as a full-WDM EGN implementation or an SSFM simulator.

### `GN_integral_math_verification.ipynb`

GN 적분 엔진의 수식과 수치 구현을 독립적으로 sanity-check하기 위한 노트북입니다. 단위 일관성, 적분 대칭성, limiting behavior, NLI의 `P^3` power law, Sobol QMC convergence 등을 작은 독립 계산으로 확인합니다.

**English:** This notebook performs independent sanity checks of the GN-integral mathematics and numerics, including unit consistency, integral symmetry, limiting behavior, the NLI `P^3` power law, and Sobol-QMC convergence.

이 파일은 API 사용법을 설명하는 tutorial보다는 수학적 검증용입니다.

**English:** It is primarily a mathematical/numerical verification notebook rather than an API tutorial.

### `GN_integral_usage_guide.ipynb`

범용 GN 엔진을 실제로 사용하는 방법을 설명하는 단계별 guide입니다. 입력/출력, SMF 및 G.654.E 예제, span sweep, launch-power sweep, self-test, 모델의 지원 범위와 제한사항을 확인할 수 있습니다.

**English:** Step-by-step usage guide for the general GN engine, including inputs/outputs, SMF and G.654.E examples, span sweeps, launch-power sweeps, self-tests, and supported-model limitations.

### `GN_modulation_Fig123_validation_colab.ipynb`

GN 적분 + modulation/system-performance 경로를 논문의 Fig. 1–3 조건에서 검증하는 핵심 Colab report입니다. PDF 벡터 그래프에서 추출한 paper reference와 코드 계산값을 같은 그래프와 표에서 비교합니다.

**English:** This is the main Colab validation report for the GN-integral plus modulation/system-performance path. It compares code results with paper references extracted from the PDF vector plots under the Fig. 1–3 test conditions.

Fig. 1–2에서는 paper GN NLI 곡선의 오차를 평가하고, Fig. 3에서는 modulation, channel spacing, fiber 종류에 따른 maximum passing span을 비교합니다.

**English:** Figs. 1–2 evaluate error against paper GN NLI curves, while Fig. 3 compares the maximum passing span as a function of modulation, channel spacing, and fiber type.

### `paper_fig12_reference.json`

논문 Fig. 1–2에서 벡터 좌표 기반으로 추출한 기준 데이터를 저장합니다. curve fitting으로 생성한 값이 아니라 논문 그래프의 축과 vector line/marker를 직접 판독한 reference입니다.

**English:** Stores reference data extracted from the vector coordinates of paper Figs. 1–2. These are plot-derived references rather than fitted values.

### `paper_fig3_reference.json`

논문 Fig. 3의 maximum-reach/reference 값을 저장합니다.

**English:** Stores the paper-derived reference data used for the Fig. 3 maximum-reach comparison.

### `fig12_results_avg.json`

Fig. 1–2를 고정 입력 및 여러 Sobol seed로 실행한 코드 결과의 cache입니다. 매번 장시간 재계산하지 않고 validation report를 빠르게 확인할 수 있도록 사용됩니다.

**English:** Cache of Fig. 1–2 code results generated with fixed inputs and multiple Sobol seeds. It allows the validation report to be viewed without recomputing every integral.

### `fig3_code_gn_seeds.json`

Fig. 3 조건에서 GN 코드로 계산한 maximum-reach 결과와 seed별 계산값을 저장하는 cache입니다.

**English:** Cache containing GN-code maximum-reach results and seed-dependent calculations for the Fig. 3 test cases.

### `readme.md`

이 하위 폴더의 검증 방법, 수치 결과, 논문 조건, 구현 범위와 한계를 자세히 설명하는 보고서입니다.

**English:** Detailed report describing the validation method, numerical results, paper conditions, implementation scope, and limitations of the `EGNvsGN` work.

---

## 5. `EGNvsGN` 성능 결과 / Performance results

저장된 validation 결과에서 Fig. 1–2의 paper GN 곡선 대비 전체 평균 절대오차는 약 **0.0395 dB**, 평균 선형 NLI 상대오차는 약 **0.9052%**입니다. 따라서 범용 GN 적분 경로는 해당 논문 조건의 GN baseline을 높은 일관성으로 재현합니다.

**English:** In the stored validation results, the overall mean absolute error against the paper GN curves of Figs. 1–2 is approximately **0.0395 dB**, with an average linear-NLI relative error of approximately **0.9052%**. The general GN-integral path therefore reproduces the paper GN baseline with strong consistency under these tested conditions.

Fig. 3 paper-GN maximum-reach 비교에서는 QPSK 평균 절대오차 약 **0.872 span**, 16QAM 약 **0.569 span**이 보고되었습니다.

**English:** For the Fig. 3 paper-GN maximum-reach comparison, the reported mean absolute errors are approximately **0.872 spans** for QPSK and **0.569 spans** for 16QAM.

논문 SIM 곡선과의 차이는 더 크며, 이는 현재 코드가 GN + AWGN BER approximation인 반면 논문의 SIM 결과가 SSFM 기반이기 때문입니다. 따라서 SIM과의 차이를 GN 코드의 단순 구현 오류로 해석하면 안 됩니다.

**English:** Differences relative to the paper SIM curves are larger because the current code combines GN modeling with an AWGN BER approximation, whereas the paper SIM curves are based on split-step simulation. Those differences should therefore not be interpreted simply as implementation error in the GN engine.

---

# 6. 두 폴더를 언제 사용할 것인가 / Which folder should be used?

GN 적분 엔진 자체의 정확도, 사용법, 변조별 GSNR/BER/capacity 계산을 확인하려면 `EGNvsGN`을 우선 참고하십시오.

**English:** Use `EGNvsGN` first when evaluating the accuracy, usage, modulation-dependent GSNR/BER/capacity behavior, or general applicability of the GN-integral engine.

SCI뿐 아니라 XCI/MCI까지 포함한 Carena full-EGN 수식 구현과 그 numerical convergence를 연구하려면 `EGN_Model_Validation`을 참고하십시오.

**English:** Use `EGN_Model_Validation` when studying the Carena full-EGN formulation, including SCI, XCI, MCI, and the numerical-convergence behavior of the implemented correction integrals.

```text
General GN accuracy / usage / capacity
        → EGNvsGN

SCI/XCI/MCI EGN algorithm validation
        → EGN_Model_Validation
```

---

# 7. 참고 논문 / References

### A. Carena et al., “EGN model of non-linear fiber propagation,” Optics Express 22 (2014), DOI: 10.1364/OE.22.016335

`EGN_Model_Validation/final_EGN.py`의 SCI/XCI/MCI non-Gaussian correction 구조와 Fig. 1/3/6/8 validation의 직접적인 기준 논문입니다.

**English:** This is the direct reference for the SCI/XCI/MCI non-Gaussian correction structure implemented in `final_EGN.py` and for the Fig. 1/3/6/8 validation campaign.

### P. Poggiolini et al., “A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems,” arXiv:1209.0394

GN baseline의 수치 적분 구조와 coherent optical transmission에서의 GN-model derivation을 위한 핵심 참고 문헌입니다.

**English:** Core reference for the numerical GN baseline and the derivation of the GN Model for nonlinear interference in coherent optical transmission systems.

### P. Poggiolini et al., “A Simple and Accurate Closed-Form EGN Model Formula”

`EGNvsGN` 폴더의 Fig. 1–3 validation에서 GN/EGN/SIM 비교 기준으로 사용됩니다.

**English:** Used as the paper reference for the Fig. 1–3 GN/EGN/SIM comparisons in the `EGNvsGN` directory.

### R. Dar et al., “Properties of nonlinear noise in long, dispersion-uncompensated fiber links,” Optics Express 21 (2013), DOI: 10.1364/OE.21.025685

modulation-dependent nonlinear noise와 normalized higher-order moment correction의 물리적 해석을 확인하기 위한 보조 참고 문헌입니다.

**English:** Supplementary reference for the physical interpretation of modulation-dependent nonlinear noise and normalized higher-order-moment correction terms.

---

# 8. 사용 시 주의사항 / Important notes

1. `EGN_Model_Validation`의 full-EGN XCI/MCI 결과는 현재 수치 적분 convergence 개선이 필요하므로 고객 사양 보증이나 최종 설계 sign-off용 기준값으로 사용하지 마십시오.  
   **English:** Do not use the current full-EGN XCI/MCI results for customer guarantees or final design sign-off until numerical convergence is improved.
2. `EGNvsGN`의 높은 GN 재현 정확도는 검증한 논문 조건에서의 결과이며 모든 fiber/link 조건에서 동일한 오차를 보장한다는 의미는 아닙니다.  
   **English:** The strong GN agreement reported in `EGNvsGN` applies to the validated paper conditions and does not guarantee the same error for every fiber or link configuration.
3. 새로운 fiber, 저분산 fiber 또는 Raman-heavy link에서는 independent SSFM, VPIphotonics 또는 별도 trusted model과의 교차 검증을 권장합니다.  
   **English:** For new fibers, low-dispersion fibers, or Raman-heavy links, independent cross-validation with SSFM, VPIphotonics, or another trusted model is recommended.
4. paper curve 데이터는 PDF vector plot에서 추출한 값이 포함되어 있으므로 원 논문의 raw numerical dataset과 완전히 동일한 정밀도를 갖는 것은 아닙니다.  
   **English:** Some paper-reference data were extracted from PDF vector plots and therefore do not have the same precision as an original raw numerical dataset from the authors.
