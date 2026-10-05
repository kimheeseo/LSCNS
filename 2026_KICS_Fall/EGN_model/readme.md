# EGN_model — Full EGN Solver, Validation, and GN/EGN Comparison

이 폴더는 Enhanced Gaussian Noise(EGN) 모델의 구현, 문헌 검증, 그리고 GN/EGN 비교 자료를 정리한 연구용 공간입니다.

현재 이 폴더에서 **새로운 Full-EGN 계산에 가장 우선적으로 사용할 파일은 [`EGN_adaptive.py`](./EGN_adaptive.py)** 입니다. `EGN_Model_Validation/`은 이전 full-EGN 구현과 Carena (2014) 검증 자료를 보존하고 있으며, `EGNvsGN/`은 범용 GN 적분 엔진의 수학적 검증과 paper-GN 재현을 담당합니다.

**Eng:** This directory contains the Full-EGN research solver, literature-validation material, and GN/EGN comparison studies. For new Full-EGN calculations, the recommended entry point is [`EGN_adaptive.py`](./EGN_adaptive.py). `EGN_Model_Validation/` preserves the earlier full-EGN implementation and Carena-validation artifacts, while `EGNvsGN/` focuses on validation of the general GN integral path.

---

## 1. 전체 구조 / Directory structure

```text
EGN_model/
├─ EGN_adaptive.py                  # recommended current Full-EGN solver
│
├─ EGN_Model_Validation/
│  ├─ final_EGN.py                  # earlier/reference Full-EGN implementation
│  ├─ final_EGN_Carena_Fig1_3_6_8_Validation_Colab.ipynb
│  ├─ README_final_EGN_validation.md
│  ├─ final_EGN_validation_summary.csv
│  ├─ final_EGN_paper_vs_code_detail.csv
│  ├─ Fig1_SMF_paper_vs_code.png
│  ├─ Fig3_SMF_paper_vs_code.png
│  ├─ Fig6_SMF_paper_vs_code.png
│  ├─ Fig8_SMF_paper_vs_code.png
│  └─ validation assets / bundle
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

> **Note:** `EGNvsGN/`에는 현재 독립 실행용 `.py` 파일이 직접 들어 있지 않습니다. 이 폴더의 notebook들이 상위 폴더의 `gn_integral_general.py`와 `gn_integral_general_modulation.py` 경로를 검증하는 구조입니다.
>
> **Eng:** `EGNvsGN/` currently contains notebooks/data rather than a standalone Python engine. The notebooks validate the parent-level GN engine and modulation layer.

---

# 2. `EGN_adaptive.py` — 권장 Full-EGN 모델 / Recommended Full-EGN model

## 목적 / Purpose

`EGN_adaptive.py`는 A. Carena et al., **“EGN model of non-linear fiber propagation”** (Optics Express, 2014)의 EGN formulation을 기반으로, GN baseline에 modulation-dependent non-Gaussian correction을 추가하여 다음 세 성분을 계산하는 연구용 Full-EGN solver입니다.

\[
G_{\mathrm{NLI}}^{\mathrm{EGN}}
=
G_{\mathrm{SCI}}^{\mathrm{EGN}}
+
G_{\mathrm{XCI}}^{\mathrm{EGN}}
+
G_{\mathrm{MCI}}^{\mathrm{EGN}}
\]

- **SCI** — Self-Channel Interference
- **XCI** — Cross-Channel Interference
- **MCI** — Multi-Channel Interference

**Eng:** `EGN_adaptive.py` is the recommended Full-EGN research solver. It evaluates GN baseline contributions together with modulation-dependent SCI, XCI, and MCI correction terms following Carena et al. (2014).

### 이론적 기준 / Theoretical basis

주요 기준은 Carena (2014)의 다음 formulation입니다.

- SCI: Eqs. (5)–(12)
- XCI: Eq. (18), Appendix A
- MCI: Appendix B
- numerical reduction/factorization: Appendix C

코드는 paper curve에 맞추기 위한 scale factor, empirical offset 또는 target-value fitting을 물리 계산에 사용하지 않습니다.

**Eng:** The physics calculation does not use paper-target scale factors, empirical offsets, or curve fitting.

---

## 2.1 왜 `adaptive`인가? / Why the adaptive solver?

초기 full-EGN 구현에서는 긴 coherent link에서 narrow phase-matched region 때문에 fixed/global quadrature가 적분점을 놓치거나 quadrature order에 민감해질 수 있었습니다.

`EGN_adaptive.py`는 이를 줄이기 위해 다음 numerical strategy를 사용합니다.

- exact linear-frequency inner primitives
- support-boundary splitting
- stationary-point / phase-aware splitting
- adaptive outer quadrature
- explicit positive/negative MCI region treatment
- independent frequency-resolution convergence check
- independent receiver-integration convergence check

**Eng:** The adaptive solver replaces a single global fixed-order integration strategy with phase-aware domain splitting, adaptive outer quadrature, and independent convergence checks for the frequency and receiver integrations.

---

## 2.2 주요 API / Main API

### `EGNFullOptions`

Full-EGN precision solver의 numerical tolerance와 convergence check를 제어합니다.

대표 옵션:

- `receiver_points`
- `max_receiver_points`
- `panel_order`
- `phase_step_rad`
- `quadrature_rtol`
- `convergence_rtol`
- `verify_convergence`
- `strict_convergence`

`verify_convergence=True`일 때 frequency/panel refinement와 receiver refinement를 독립적으로 확인합니다. `strict_convergence=True`에서는 요구 tolerance가 입증되지 않으면 `EGNConvergenceError`를 발생시켜 조용히 결과를 통과시키지 않습니다.

**Eng:** `EGNFullOptions` controls adaptive integration and numerical-convergence verification. With strict checking enabled, an unconverged result raises an explicit error rather than being silently accepted.

### `EGNBreakdown`

결과를 다음 성분으로 분해하여 제공합니다.

- `gn_total_W`
- `sci_gn_W`, `xci_gn_W`, `mci_gn_W`
- `sci_correction_W`, `xci_correction_W`, `mci_correction_W`
- `total_egn_W`
- `egn_sci_W`, `egn_xci_W`, `egn_mci_W`, `egn_xmci_W`
- GN 대비 EGN ratio / dB difference
- convergence diagnostics

### `egn_span_sweep(...)`

동일 span의 개수를 여러 값으로 바꾸면서 GN/EGN 성분을 계산하는 대표 public API입니다. span-by-span NLI accumulation이나 Carena Fig. 1/3/6/8 형태의 비교에 적합합니다.

**Eng:** `egn_span_sweep(...)` is the main precision API for evaluating several span counts using the same physical system and adaptive numerical primitives.

---

## 2.3 지원하는 precision Full-EGN 범위 / Supported precision scope

현재 precision Full-EGN path는 다음 조건을 대상으로 합니다.

- dual-polarization convention
- coherent accumulation
- rectangular / zero-rolloff spectrum
- equal symbol rates
- non-overlapping WDM channels
- identical, loss-compensated lumped-EDFA spans
- `beta2`-dominated dispersion (`beta3=0` in the precision EGN path)
- MCI 계산 시 odd, symmetric, equally-spaced WDM comb
- MCI 계산 시 equal channel powers and common modulation distribution
- central channel as CUT for Appendix-B MCI calculation

SCI와 XCI는 1/2-channel 구성에서도 계산할 수 있지만, 3채널 이상에서 Appendix-B MCI를 사용하려면 위의 symmetric-comb 조건이 필요합니다.

**Eng:** The precision Full-EGN path intentionally enforces the validated assumptions of the implemented Carena formulation instead of silently applying it outside its supported scope.

---

## 2.4 Full-EGN path와 legacy GN path의 차이

`EGN_adaptive.py` 내부에는 기존 GN API의 보다 넓은 기능도 유지되어 있습니다. 예를 들어 `beta3`, distributed gain, heterogeneous spans 등은 GN 경로에서 사용할 수 있습니다.

그러나 이러한 조건은 precision Full-EGN path에서 자동으로 일반화되지 않습니다. 검증되지 않은 profile은 Full-EGN 계산에서 reject됩니다.

**Eng:** The file retains broader legacy GN capabilities, but the precision Full-EGN solver deliberately rejects unvalidated profiles such as heterogeneous spans, distributed gain, or nonzero `beta3`.

---

## 2.5 간단한 실행 예 / Minimal example

```python
from EGN_adaptive import WDMSystem, Span, EGNFullOptions, egn_span_sweep

system = WDMSystem.equispaced(
    n_channels=3,
    spacing_GHz=33.6,
    baud_GBd=32.0,
    power_dBm=0.0,
    pulse_shape="rect",
    rolloff=0.0,
)

span = Span(
    length_km=100.0,
    alpha_db_per_km=0.22,
    gamma_W_inv_km=1.3,
    D_ps_nm_km=16.7,
)

options = EGNFullOptions(
    verify_convergence=True,
    strict_convergence=True,
)

results = egn_span_sweep(
    system=system,
    span=span,
    span_counts=[1, 2, 5, 10, 20, 50],
    cut_index=1,
    modulation="QPSK",
    full_options=options,
)

for nspan, result in results.items():
    print(
        nspan,
        result.egn_sci_W,
        result.egn_xci_W,
        result.egn_mci_W,
        result.total_egn_W,
        result.diagnostics.converged,
    )
```

---

## 2.6 정확도 해석 / How to interpret accuracy

`diagnostics.converged=True`는 **해당 numerical refinement에서 계산값이 설정한 tolerance 이내로 안정화되었다는 의미**입니다.

이는 다음을 자동으로 보장하지 않습니다.

- paper curve 대비 항상 `<3%`
- SSFM 대비 항상 특정 오차 이내
- 임의의 실제 실험 조건에서 동일 정확도

따라서 새로운 시스템 조건에서는 가능하면 paper reference, independent SSFM, 실험값 또는 상용 optical-system simulator와 일부 reference point를 교차 검증한 후 parameter sweep에 사용하는 것을 권장합니다.

**Eng:** Numerical convergence is not the same as physical-model validation. A converged integral does not guarantee a universal error bound versus SSFM or experiment.

---

# 3. `EGN_Model_Validation/`

## 역할 / Role

이 폴더는 Carena (2014)의 Fig. 1, 3, 6, 8 조건을 이용하여 이전/reference full-EGN 구현인 `final_EGN.py`를 검증한 자료를 보관합니다.

**Eng:** This folder preserves the earlier/reference Full-EGN implementation and the associated Carena Fig. 1/3/6/8 validation artifacts.

### Python file: `final_EGN.py`

Poggiolini-style GN integral을 baseline으로 사용하고 Carena의 SCI/XCI/MCI non-Gaussian correction을 추가한 이전 full-EGN 구현입니다.

이 파일은 현재 `EGN_adaptive.py`가 개선하려고 한 numerical-convergence 문제를 확인하는 데 중요한 **reference/validation implementation**으로 남겨두는 것이 적절합니다.

**Eng:** `final_EGN.py` is retained as a reference implementation and validation baseline. It is useful for understanding the fixed/global quadrature sensitivity that motivated the phase-aware adaptive solver.

### Validation notebook

`final_EGN_Carena_Fig1_3_6_8_Validation_Colab.ipynb`

- Fig. 1: SCI
- Fig. 3: 3-channel XCI only
- Fig. 6: 3-channel XMCI = XCI + MCI
- Fig. 8: 9-channel XMCI = XCI + MCI

paper-derived reference와 code result를 graph/table/error metric으로 비교합니다.

### Validation data and figures

- `README_final_EGN_validation.md`: 상세 validation report
- `final_EGN_validation_summary.csv`: figure/fiber별 summary
- `final_EGN_paper_vs_code_detail.csv`: point-by-point paper/code/error
- `Fig1_SMF_paper_vs_code.png`: SCI comparison
- `Fig3_SMF_paper_vs_code.png`: XCI-only comparison
- `Fig6_SMF_paper_vs_code.png`: 3-channel XMCI comparison
- `Fig8_SMF_paper_vs_code.png`: 9-channel XMCI comparison
- validation bundle/assets: 재현 및 보관용 결과물

### 해석 / Interpretation

`final_EGN.py`의 GN baseline과 SCI-EGN은 비교적 안정적인 반면, 기존 XCI/MCI correction path는 quadrature-order sensitivity가 확인되었습니다. 이 때문에 새로운 Full-EGN 연구 계산은 root의 `EGN_adaptive.py`를 우선 사용하는 것을 권장합니다.

**Eng:** The earlier implementation is valuable as a validation/reference baseline, but new Full-EGN studies should preferentially use the adaptive solver.

---

# 4. `EGNvsGN/`

## 역할 / Role

이 폴더는 Full-EGN solver 자체보다는 `gn_integral_general.py`와 `gn_integral_general_modulation.py`의 **GN baseline, 수학적 일관성, 사용법, paper-GN 재현성**을 검증하기 위한 자료입니다.

**Eng:** This folder validates the general GN engine and modulation/system-performance path rather than serving as a standalone Full-EGN implementation.

### Python file 유무 / Python files

현재 `EGNvsGN/` 폴더 안에는 독립적인 `.py` 모델 파일이 없습니다.

검증 대상 Python engine은 상위 폴더의:

- `../../gn_integral_general.py`
- `../../gn_integral_general_modulation.py`

입니다.

### `GN_integral_math_verification.ipynb`

GN 적분식의 unit consistency, symmetry, limiting behavior, NLI `P^3` law 및 Sobol-QMC convergence를 독립적으로 sanity-check합니다.

### `GN_integral_usage_guide.ipynb`

GN engine의 입력/출력과 SMF/G.654.E 예제, span sweep, launch-power sweep, self-test 및 모델의 지원 범위를 설명합니다.

### `GN_modulation_Fig123_validation_colab.ipynb`

GN integral + modulation/system layer를 관련 논문의 Fig. 1–3 조건에서 검증하는 report입니다.

- Fig. 1–2: paper GN NLI curve 비교
- Fig. 3: modulation/channel spacing/fiber에 따른 maximum passing span 비교
- published EGN/SIM curve는 reference로 함께 표시하지만 현재 GN engine이 이를 직접 구현했다고 주장하지 않음

### JSON reference/result files

- `paper_fig12_reference.json`: paper Fig. 1–2 reference
- `paper_fig3_reference.json`: paper Fig. 3 reference
- `fig12_results_avg.json`: Fig. 1–2 code result cache
- `fig3_code_gn_seeds.json`: Fig. 3 seed-based result cache

**Eng:** These JSON files separate paper-derived references from code-generated results, which improves reproducibility and prevents hidden fitting.

---

# 5. 모델 선택 가이드 / Model-selection guide

| 목적 | 사용할 파일 | 설명 |
|---|---|---|
| 범용 GN NLI | `../gn_integral_general.py` | heterogeneous spans, flexible PSD, beta2/beta3 등 가장 넓은 GN 범위 |
| GN + ASE/GSNR/BER/capacity | `../gn_integral_general_modulation.py` | system-performance layer |
| SCI-only EGN correction | modulation layer의 `nli_model="egn_sci"` | Full EGN 아님 |
| **Full EGN SCI+XCI+MCI** | **`EGN_adaptive.py`** | 현재 권장 research solver |
| 과거 full-EGN validation 재현 | `EGN_Model_Validation/final_EGN.py` | reference/validation implementation |
| GN 수학/논문 검증 | `EGNvsGN/` notebooks | paper-GN 및 QMC 검증 |
| SSFM/NLSE propagation | 현재 미구현 | 별도 simulator 필요 |

---

# 6. 권장 사용 범위 / Recommended use

`EGN_adaptive.py`는 다음과 같은 연구에 적합합니다.

- fiber type에 따른 NLI 비교
- span length / number-of-spans 변화
- channel spacing 변화
- WDM channel-count 변화
- modulation-dependent EGN correction
- SCI/XCI/MCI breakdown
- GN vs EGN comparison
- experimental/SSFM study 전 사전 NLI prediction

반면 다음 조건에서는 별도 검증 또는 모델 확장이 필요합니다.

- non-zero RRC roll-off의 Full-EGN
- arbitrary heterogeneous Full-EGN spans
- distributed Raman Full-EGN
- mid-link add/drop / ROADM filtering history
- PMD / PDL
- laser phase noise
- transceiver implementation penalties
- HCF/MCF 고유 물리효과
- core/mode coupling

---

# 7. Dependencies

```text
Python >= 3.10
numpy
scipy
```

---

# 8. References

1. A. Carena, G. Bosco, V. Curri, Y. Jiang, P. Poggiolini, F. Forghieri, **“EGN model of non-linear fiber propagation,”** Optics Express 22, 16335–16362 (2014), DOI: 10.1364/OE.22.016335.
2. P. Poggiolini et al., **“A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems,”** arXiv:1209.0394.
3. P. Poggiolini, **“The GN Model of Non-Linear Propagation in Uncompensated Coherent Optical Systems,”** Journal of Lightwave Technology 30(24), 3857–3879 (2012).

---

## Summary

- **새로운 Full-EGN 계산:** `EGN_adaptive.py`
- **이전 Full-EGN 구현/검증 자료:** `EGN_Model_Validation/`
- **범용 GN engine 검증 및 사용 가이드:** `EGNvsGN/`

**Eng:** Use `EGN_adaptive.py` for new Full-EGN research calculations, `EGN_Model_Validation/` for the earlier implementation and Carena-validation artifacts, and `EGNvsGN/` for GN-engine verification and usage studies.
