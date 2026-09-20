# 2026_KICS_Fall_10 — Numerical GN Model and System-Performance Tools

이 폴더는 코히어런트 광전송 시스템의 비선형 간섭(Nonlinear Interference, NLI)을 **수치 적분 기반 Gaussian Noise (GN) Model**로 계산하고, 그 결과를 변조 방식·ASE·GSNR·BER·전송률/용량 분석으로 확장하기 위한 연구용 코드입니다.

**Eng:** This directory contains a numerical Gaussian Noise (GN) model engine for nonlinear-interference estimation in coherent optical transmission systems, together with a modulation/system-performance layer for ASE, GSNR, BER, rate, and capacity analysis.

> **Important:** `gn_integral_general.py`와 `gn_integral_general_modulation.py`의 기본 경로는 **GN Model**입니다. Full-WDM EGN(SCI+XCI+MCI correction)을 사용하려면 [`EGN_model/EGN_adaptive.py`](./EGN_model/EGN_adaptive.py)를 사용하십시오.
>
> **Eng:** The default path in the two files in this directory is the **GN Model**. For a Full-WDM EGN calculation including SCI/XCI/MCI corrections, use [`EGN_model/EGN_adaptive.py`](./EGN_model/EGN_adaptive.py).

---

## 1. 파일 구성 / Main files

| 파일 | 권장 용도 | 핵심 역할 |
|---|---|---|
| [`gn_integral_general.py`](./gn_integral_general.py) | GN 기반 NLI 계산 | 전체 WDM PSD와 링크 물리를 직접 적분하는 핵심 엔진 |
| [`gn_integral_general_modulation.py`](./gn_integral_general_modulation.py) | 시스템 성능 분석 | 변조 moments, ASE, GSNR, BER, rate/capacity, launch-power sweep |
| [`EGN_model/EGN_adaptive.py`](./EGN_model/EGN_adaptive.py) | Full-EGN 연구 계산 | Carena (2014) 기반 SCI+XCI+MCI correction 및 adaptive convergence check |

**Eng:** `gn_integral_general.py` is the physical GN/NLI engine, `gn_integral_general_modulation.py` is the modulation and performance layer, and `EGN_model/EGN_adaptive.py` is the recommended Full-EGN research solver.

---

## 2. `gn_integral_general.py`

`gn_integral_general.py`는 Poggiolini 계열 GN formulation을 기반으로, 전체 WDM power spectral density를 2차원 주파수 영역에서 직접 적분하여 NLI PSD와 선택 채널의 NLI power를 계산합니다.

**Eng:** This is the core numerical GN engine. It directly evaluates the two-dimensional frequency-domain GN integral over the complete WDM PSD and returns the NLI PSD or the integrated NLI power of a selected channel.

### 주요 기능 / Main capabilities

- 전체 WDM PSD 적분. GN 관점의 SCI/XCI/MCI interaction region이 전체 적분 안에 자연스럽게 포함됩니다.  
  **Eng:** Full-WDM PSD integration; interaction regions corresponding to SCI/XCI/MCI are naturally included in the total **GN** integral.
- 채널별 서로 다른 power, bandwidth 및 비균일 channel spacing 지원.  
  **Eng:** Unequal channel powers/bandwidths and irregular channel spacing.
- rectangular, raised-cosine, RRC-power, custom PSD 지원.  
  **Eng:** Rectangular, raised-cosine, RRC-power, and user-defined PSDs.
- `beta2` + `beta3` phase mismatch.  
  **Eng:** Phase mismatch including both `beta2` and `beta3`.
- span별 length, attenuation, dispersion, `gamma`, gain이 다른 heterogeneous link.  
  **Eng:** Heterogeneous links with span-dependent fiber/link parameters.
- coherent 또는 incoherent span accumulation.  
  **Eng:** Selectable coherent or incoherent span accumulation.
- lumped EDFA, ideal distributed gain, simplified backward-Raman profile, custom gain profile.  
  **Eng:** Lumped EDFA and several numerical distributed-gain profiles.
- scrambled Sobol quasi-Monte-Carlo(QMC) 기반 2-D 적분.  
  **Eng:** Scrambled Sobol QMC for the 2-D frequency integral.

### 중요한 convention / Important conventions

- frequency: THz (= 1/ps)
- `beta2`: ps²/km
- `beta3`: ps³/km
- distance: km
- `gamma`: 1/(W·km)
- `alpha_db_per_km`: optical **power** attenuation [dB/km]
- channel power / PSD: total dual-polarization quantity
- DP GN coefficient: `16/27`

### GN과 제한적 EGN 옵션의 차이 / GN vs optional SCI-EGN

기본값은 Gaussian-signal assumption을 사용하는 순수 GN Model입니다. `egn_mu4`와 `egn_mu6`를 명시하면 Carena et al. (2014)의 rectangular-spectrum **SCI correction**을 선택적으로 사용할 수 있습니다.

그러나 이 옵션은 XCI/MCI에 대한 비가우시안 EGN correction을 적용하지 않습니다. 따라서 이 파일만으로 계산한 `egn_sci` 결과를 **Full-WDM EGN**으로 해석하면 안 됩니다.

**Eng:** Optional `egn_mu4`/`egn_mu6` parameters enable the Carena rectangular-spectrum SCI correction. XCI/MCI remain on the GN path, so this option is not a Full-WDM EGN implementation.

---

## 3. `gn_integral_general_modulation.py`

이 파일은 `gn_integral_general.py`의 NLI 계산을 그대로 사용하면서 변조 방식과 시스템 성능 계산을 추가하는 상위 레이어입니다.

**Eng:** This module is a higher-level modulation and system-performance layer built on top of `gn_integral_general.py`; it delegates the NLI integration to the core engine rather than re-implementing it.

### 지원 변조 / Supported modulation formats

BPSK, QPSK, 8QAM, 16QAM, 32QAM, 64QAM, 256QAM

각 explicit constellation에서 다음 normalized moments를 직접 계산합니다.

\[
\Phi = \mu_4 - 2,
\qquad
\Psi = \mu_6 - 9\mu_4 + 12
\]

따라서 modulation-dependent moment를 paper curve에 맞춘 empirical fitting 값으로 사용하지 않습니다. 다만 8QAM과 32QAM은 코드에 명시된 특정 constellation geometry를 사용하므로 다른 mapping을 사용할 경우 moments를 다시 계산해야 합니다.

**Eng:** Modulation moments are calculated from explicit constellation coordinates rather than fitted correction factors. The built-in 8QAM/32QAM geometries are modeling choices and should be replaced if another constellation is intended.

### 시스템 성능 계산 / Performance calculations

- EDFA ASE noise accumulation
- selected CUT NLI power
- `SNR_ASE`, `SNR_NLI`, optional transceiver SNR, GSNR
- approximate AWGN BER
- launch-power vs. GSNR sweep
- dual-polarization gross/net line rate
- Shannon-gap capacity estimate

`nli_model="gn"`이 기본값입니다. `nli_model="egn_sci"`는 modulation moments를 core engine으로 전달하여 SCI만 EGN으로 보정합니다. XCI/MCI는 GN으로 유지됩니다.

**Eng:** `nli_model="gn"` is the default. `nli_model="egn_sci"` adds modulation-dependent EGN correction to SCI only; XCI/MCI remain GN.

---

## 4. 코드 관계 / Software architecture

```text
WDM / fiber / span inputs
          │
          ▼
gn_integral_general.py
  ├─ WDM PSD
  ├─ link function / phase mismatch
  ├─ coherent or incoherent accumulation
  └─ numerical GN integral
          │
          ├──────────────► NLI PSD / NLI power
          │
          ▼
gn_integral_general_modulation.py
  ├─ modulation moments
  ├─ ASE
  ├─ GSNR / BER
  ├─ rate / capacity
  └─ launch-power sweep

For Full EGN:
EGN_model/EGN_adaptive.py
  └─ GN baseline + SCI correction + XCI correction + MCI correction
```

---

## 5. 간단한 사용 예 / Minimal example

```python
from gn_integral_general import WDMSystem, Span
from gn_integral_general_modulation import launch_power_vs_gsnr

system = WDMSystem.equispaced(
    n_channels=9,
    spacing_GHz=50.0,
    baud_GBd=32.0,
    power_dBm=0.0,
    pulse_shape="rect",
)

spans = [
    Span(
        length_km=80.0,
        alpha_db_per_km=0.20,
        gamma_W_inv_km=1.3,
        D_ps_nm_km=17.0,
        noise_figure_db=5.0,
    )
    for _ in range(10)
]

result = launch_power_vs_gsnr(
    base_system=system,
    spans=spans,
    cut_index=4,
    launch_power_dBm=[-4, -2, 0, 2, 4],
    modulation="QPSK",
    nli_model="gn",
)
```

고정된 링크에서 launch power에 따른 NLI/GSNR 경향을 확인하거나, fiber parameter(`alpha`, `D`, `gamma`)와 span 수를 변경하여 링크 성능을 비교하는 용도로 사용할 수 있습니다.

**Eng:** The same API can be used for launch-power sweeps, span-count studies, and comparisons among different fiber parameters.

---

## 6. 어떤 모델을 사용해야 하는가? / Which model should I use?

| 목적 | 권장 경로 | 현재 위치 |
|---|---|---|
| 범용 WDM GN NLI | `gn_integral_general.py` | 가장 넓은 링크/PSD 범위 |
| ASE + GSNR + BER + capacity | `gn_integral_general_modulation.py`, `nli_model="gn"` | 시스템 성능 분석용 |
| modulation-dependent SCI 보정 | `nli_model="egn_sci"` | SCI only, Full EGN 아님 |
| SCI + XCI + MCI Full-EGN | `EGN_model/EGN_adaptive.py` | 검증된 precision scope 내 research solver |
| SSFM/NLSE waveform propagation | 현재 미구현 | 별도 SSFM 또는 상용 simulator 필요 |

실무/연구에서는 GN과 EGN 중 하나를 고르는 방식보다 다음과 같은 계층형 사용을 권장합니다.

```text
GN screening / parameter sweep
        ↓
EGN refinement at selected operating points
        ↓
SSFM / VPI / experiment cross-validation
```

GN은 넓은 설계 공간을 빠르게 탐색하는 데 적합하고, EGN은 변조 의존성과 SCI/XCI/MCI의 비가우시안 보정을 정밀하게 확인하는 데 적합합니다.

---

## 7. 레퍼런스 논문 구현 충실도 / Fidelity to the reference formulations

### 7.1 GN path

GN core는 Poggiolini 계열 formulation의 핵심 항목을 그대로 구현합니다.

- power attenuation [dB/km] → field attenuation 변환
- `beta2` / `beta3` 기반 four-wave-mixing phase mismatch
- single-span nonlinear source integral
- coherent / incoherent span accumulation
- multi-span phase history
- full-WDM 2-D frequency integration
- exact first-order cubic NLI scaling: `P_NLI ∝ P_ch^3`

이 구현은 특정 paper curve에 맞춘 fitting model이 아니라, **논문 수식을 범용 numerical solver로 확장한 형태**입니다. 따라서 irregular spacing, unequal channel powers, custom PSD, unequal spans, beta3, heterogeneous span parameters까지 GN 경로에서 처리할 수 있습니다.

`gn_integral_general_modulation.py`는 GN 수식을 다시 구현하지 않고, 이 GN core 위에 ASE, GSNR, BER, modulation, throughput/capacity 계산을 추가합니다.

### 7.2 EGN path

`EGN_model/EGN_adaptive.py`는 Carena et al. (2014)의 Full-EGN formulation을 기준으로 다음을 구현합니다.

- SCI: Eqs. (5)–(12)
- XCI: Eq. (18), Appendix A
- MCI: Appendix B
- numerical reduction / factorization: Appendix C
- modulation moments `mu4`, `mu6`
- SCI/XCI/MCI component breakdown
- phase-aware adaptive numerical integration
- independent frequency / receiver convergence checks

물리 계산에는 paper curve에 맞추기 위한 scale factor, empirical offset, target-value fitting을 사용하지 않습니다.

다만 precision Full-EGN path는 논문의 검증된 조건을 의도적으로 따릅니다.

- rectangular / zero-rolloff spectrum
- equal symbol rate
- coherent accumulation
- identical loss-compensated lumped-EDFA spans
- beta2-dominated dispersion
- symmetric/equispaced WDM comb for Appendix-B MCI
- central CUT / equal-power assumptions where required

즉 **EGN은 GN보다 정밀하지만, 현재 검증된 적용 범위는 GN보다 좁습니다.**

---

## 8. 검증 성능 / Validation performance

### 8.1 GN — equation-level reproducibility

실행형 validation은 다음 위치에 있습니다.

- [GN_model/README.md](./GN_model/README.md)
- [GN_Model_Colab.ipynb](./GN_model/GN_Model_Colab.ipynb)
- [results/COLAB_EXECUTION_REPORT.md](./GN_model/results/COLAB_EXECUTION_REPORT.md)

Poggiolini 기반 equation-level test 결과:

| 항목 | 결과 |
|---|---:|
| Equation-level assessment | **PASS** |
| Maximum equation relative error | **1.29452 × 10⁻¹¹ %** |
| Sobol relative std @ 32,768 samples | **0.0935 %** |
| 8,192 → 32,768 sample change | **0.1441 %** |

이 결과는 GN 식 자체의 구현과 numerical integration이 매우 안정적임을 의미합니다.

### 8.2 GN — Carena 2012 Fig. 5 maximum-reach reproduction

Carena et al. JLT 2012 Fig. 5의 63개 digitized point를 이용한 end-to-end reach validation 결과:

| Metric | GN result |
|---|---:|
| Compared points | **63** |
| Overall MAPE | **8.66 %** |
| Median APE | **6.67 %** |
| Points within 5 % | **36.5 %** |
| Points within 10 % | **69.8 %** |
| Points within 15 % | **85.7 %** |
| Points within 20 % | **92.1 %** |
| Reach ≥ 1,000 km MAPE | **7.38 %** |

Fiber-level MAPE:

| Fiber | MAPE |
|---|---:|
| PSCF | 9.58 % |
| SMF | 6.95 % |
| NZDSF | 9.43 % |

Modulation-level MAPE:

| Modulation | MAPE |
|---|---:|
| BPSK | 9.31 % |
| **QPSK** | **5.20 %** |
| 8QAM | 10.11 % |
| 16QAM | 11.03 % |

최대 percentage error는 short-reach point에서 크게 나타날 수 있습니다. Fig. 5는 100-km span 단위의 maximum reach이고, reference 값도 PDF에서 digitize했기 때문에 짧은 거리에서는 span 하나 차이가 큰 percentage error로 보일 수 있습니다.

Paper-derived Fig. 3 / Fig. 5 reference는 원 저자 raw table이 아니라 PDF plot에서 추출했으며, validation data에는 약 **4–8 % digitization uncertainty**가 별도로 표시되어 있습니다.

### 8.3 EGN_adaptive — Carena 2014 analytical EGN reproduction

현재 `EGN_adaptive.py`의 직접 검증은 다음 위치에 있습니다.

- [EGN_MODEL_Validation/README.md](./EGN_model/EGN_MODEL_Validation/README.md)
- [EGN_Model_Validation_Colab.ipynb](./EGN_model/EGN_MODEL_Validation/EGN_Model_Validation_Colab.ipynb)
- [paper_vs_code_error_summary.csv](./EGN_model/EGN_MODEL_Validation/paper_vs_code_error_summary.csv)

Carena et al. Optics Express 2014의 analytical EGN curves를 Fig. 1 / 3 / 6 / 8에서 비교한 결과:

| Item | Result |
|---|---:|
| Compared figures | Fig. 1, 3, 6, 8 |
| Compared points | **72** |
| Overall MAE | **0.136 dB** |
| Overall RMSE | **0.193 dB** |
| Mean relative error in linear eta | **3.13 %** |
| Maximum absolute error | **0.745 dB** |
| Representative high-resolution repeat | **≤ 0.032 dB change** |

세부 panel별 mean relative error:

| Case | SMF | NZDSF | LS |
|---|---:|---:|---:|
| Fig. 1 SCI | 5.02 % | 2.75 % | 1.95 % |
| Fig. 3 XCI, 3 ch | 2.71 % | 2.09 % | 3.12 % |
| Fig. 6 XMCI, 3 ch | **0.74 %** | 1.54 % | 3.39 % |
| Fig. 8 XMCI, 9 ch | 2.37 % | 6.18 % | 5.75 % |

따라서 현재 검증 범위에서 `EGN_adaptive.py`는 Carena 2014 analytical EGN result를 평균적으로 약 **3.1 %** 수준으로 재현합니다.

**중요:** 이는 SSFM/실험 대비 3.13 %라는 뜻이 아닙니다. 동일한 analytical EGN reference를 얼마나 충실하게 구현했는지를 나타내는 값입니다. SSFM 또는 실험과의 독립 검증은 별도의 단계입니다.

### 8.4 GN Fig. 5에서 보인 EGN proxy 오차를 어떻게 해석해야 하는가

GN Fig. 5 비교 과정에서 EGN one-span result를 paper의 incoherent accumulation convention에 맞춰 비교한 proxy는 큰 오차를 보였습니다. 이 값은 `EGN_adaptive.py`의 native Full-EGN 정확도 점수로 사용하면 안 됩니다.

이유는 다음과 같습니다.

- Carena 2012 Fig. 5: **GN benchmark + incoherent span accumulation**
- `EGN_adaptive.py`: **coherent Full-EGN + rectangular-spectrum precision path**

따라서 EGN 자체의 구현 성능은 위의 **Carena 2014 Fig. 1/3/6/8, 72-point validation**을 기준으로 해석해야 합니다.

---

## 9. 범용성 및 권장 활용 / Generality and recommended use

### GN model

현재 GN 모델은 다음 용도로 범용적으로 사용할 수 있습니다.

- SMF / G.652.D
- G.654.E
- 일반 WDM coherent link
- launch-power optimization
- span-length / span-count sweep
- fiber parameter comparison
- attenuation / dispersion / gamma / Aeff sensitivity
- GSNR / BER / throughput / capacity screening
- repeater / span architecture preliminary study

특히 다음과 같은 engineering trend 분석에 적합합니다.

```text
Aeff ↑
  → gamma = 2π n2 / (lambda · Aeff) ↓
  → P_NLI ↓
  → GSNR ↑
```

즉 G.652.D vs G.654.E, span length, launch power, channel spacing, WDM bandwidth, repeater 수에 따른 상대 비교를 빠르게 수행하는 **general-purpose GN screening / design engine**으로 사용하는 것이 적절합니다.

현재 독립 Fig. 5 validation을 기준으로는 모든 조건에서 <5 %라고 주장하면 안 되지만, tested domain에서 대체로 **약 7–9 % 수준의 reach-prediction accuracy**, QPSK에서는 약 **5.2 % MAPE**를 보였습니다.

### EGN_adaptive

`EGN_adaptive.py`는 다음 용도로 적합합니다.

- modulation-dependent nonlinear correction
- SCI / XCI / MCI breakdown
- GN vs EGN comparison
- low-dispersion / strongly modulation-dependent cases
- selected operating-point refinement
- paper reproduction / advanced NLI research

Carena 2014 analytical EGN reproduction은 평균 **3.13 %** 수준이므로, 검증된 scope 내에서는 research-grade refinement model로 사용할 수 있습니다.

하지만 GN보다 적용 범위가 좁기 때문에 arbitrary heterogeneous link, distributed Raman, nonzero beta3, ROADM filtering history 등에서는 native precision Full-EGN result를 일반화하면 안 됩니다.

### 권장 workflow

```text
GN screening
    ↓
EGN refinement
    ↓
SSFM / VPI / experiment validation
```

예를 들어 해저 광전송 시스템에서는 GN으로 수천 개의 fiber / span / launch-power 조합을 먼저 탐색하고, 최종 후보 operating point만 EGN으로 다시 계산한 뒤 SSFM/VPI 또는 실험값으로 최종 확인하는 구조가 효율적입니다.

---

## 10. 사용 시 주의사항 / Limitations and interpretation

- GN Model은 Gaussian-signal assumption을 사용합니다.
- QMC 결과는 `sobol_power`, seed, receiver integration resolution에 대해 convergence를 확인하는 것이 좋습니다.
- distributed Raman/custom gain은 GN 경로의 numerical extension이며 실제 system sign-off 전 독립 검증이 필요합니다.
- `egn_sci`는 Full EGN이 아닙니다. SCI+XCI+MCI correction에는 `EGN_model/EGN_adaptive.py`를 사용하십시오.
- Full-EGN precision path는 rectangular/equal-baud/homogeneous coherent-link assumptions 안에서 사용해야 합니다.
- 본 repository는 SSFM simulator가 아니며 PMD, PDL, laser phase noise, DSP implementation penalty, nonlinear phase-noise dynamics 등 모든 실험 요소를 포함하지 않습니다.
- customer-facing guarantee, certification, final design sign-off에는 SSFM/VPI/experimental reference point를 함께 사용하는 것을 권장합니다.

---

## 11. Dependencies

```text
Python >= 3.10
numpy
scipy
```

Colab 또는 일반 Python 환경에서 사용할 수 있습니다.

---

## 12. References

1. P. Poggiolini et al., **“A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems”**, arXiv:1209.0394.
2. P. Poggiolini, **“The GN Model of Non-Linear Propagation in Uncompensated Coherent Optical Systems”**, Journal of Lightwave Technology 30(24), 3857–3879 (2012).
3. A. Carena et al., **“Modeling of the Impact of Nonlinear Propagation Effects in Uncompensated Optical Coherent Transmission Links”**, Journal of Lightwave Technology 30(10), 1524–1539 (2012).
4. A. Carena et al., **“EGN model of non-linear fiber propagation”**, Optics Express 22, 16335–16362 (2014), DOI: 10.1364/OE.22.016335.

---

## Summary

현재 코드의 역할은 다음처럼 구분하는 것이 가장 적절합니다.

- **`gn_integral_general.py` + `gn_integral_general_modulation.py`**  
  → 논문 수식 구현이 매우 안정적이고, 범용 WDM / link / system-performance study에 적합한 **general-purpose GN screening/design engine**  
  → Carena 2012 Fig. 5 기준 전체 MAPE **8.66 %**, long-reach subset MAPE **7.38 %**, QPSK MAPE **5.20 %**

- **`EGN_model/EGN_adaptive.py`**  
  → Carena 2014 SCI/XCI/MCI analytical EGN을 직접 구현한 **research-grade Full-EGN refinement solver**  
  → 72 reference points 기준 MAE **0.136 dB**, mean linear-eta relative error **3.13 %**

따라서 이 repository는 단순한 paper-reproduction code라기보다,

```text
GN = broad engineering design / fast screening
EGN = high-accuracy nonlinear refinement within validated scope
SSFM/VPI/experiment = final independent reference
```

라는 계층형 optical-link simulation workflow로 사용하는 것이 가장 적절합니다.
