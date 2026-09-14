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

| 목적 | 권장 경로 | 비고 |
|---|---|---|
| 범용 WDM GN NLI | `gn_integral_general.py` | 가장 넓은 링크/PSD 범위 |
| ASE + GSNR + BER + capacity | `gn_integral_general_modulation.py`, `nli_model="gn"` | 시스템 성능 분석용 |
| modulation-dependent SCI 보정 | `nli_model="egn_sci"` | **SCI only**, Full EGN 아님 |
| SCI + XCI + MCI Full-EGN | `EGN_model/EGN_adaptive.py` | 검증된 precision scope 내 사용 권장 |
| SSFM/NLSE waveform propagation | 현재 미구현 | 별도 SSFM 또는 상용 simulator 필요 |

---

## 7. 검증 / Validation

`EGN_model/EGNvsGN/GN_modulation_Fig123_validation_colab.ipynb`는 동일 GN 경로를 문헌의 GN reference curve와 비교합니다.

저장된 결과 기준 Fig. 1–2 paper-GN 곡선 대비:

- 전체 평균 절대오차: 약 **0.0395 dB**
- 평균 linear-NLI 상대오차: 약 **0.9052%**

Fig. 3 maximum-reach의 paper-GN curve 대비 평균 절대오차는 QPSK 약 **0.872 span**, 16QAM 약 **0.569 span**입니다.

이 수치는 **해당 논문 조건에 대한 검증 결과**이며 임의의 시스템에서 동일한 정확도를 보장한다는 의미가 아닙니다.

**Eng:** The stored validation results reproduce the tested paper-GN curves closely, but the reported errors apply only to the tested configurations and are not a universal accuracy guarantee.

---

## 8. 사용 시 주의사항 / Limitations and interpretation

- GN Model은 Gaussian-signal assumption을 사용합니다.  
  **Eng:** The GN Model relies on the Gaussian-signal assumption.
- QMC 결과는 `sobol_power`, seed, receiver integration resolution에 대해 convergence를 확인하는 것이 좋습니다.  
  **Eng:** High-accuracy use should include convergence checks versus QMC sample count/seed and receiver integration resolution.
- distributed Raman/custom gain은 범용 numerical extension이며 실제 시스템 sign-off 전 독립 검증이 필요합니다.  
  **Eng:** Distributed-Raman/custom gain support should be independently validated before design sign-off.
- `egn_sci`는 Full EGN이 아닙니다. Full EGN 연구는 `EGN_model/EGN_adaptive.py`를 사용하십시오.  
  **Eng:** `egn_sci` is not Full EGN; use `EGN_model/EGN_adaptive.py` for SCI+XCI+MCI correction.
- 본 코드는 SSFM simulator가 아니며 PMD, PDL, laser phase noise, DSP penalty 등 모든 실험 요소를 포함하지 않습니다.  
  **Eng:** This repository is not an SSFM simulator and does not model every experimental impairment.

---

## 9. Dependencies

```text
Python >= 3.10
numpy
scipy
```

Colab 또는 일반 Python 환경에서 사용할 수 있습니다.

---

## 10. References

1. P. Poggiolini et al., **“A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems”**, arXiv:1209.0394.
2. P. Poggiolini, **“The GN Model of Non-Linear Propagation in Uncompensated Coherent Optical Systems”**, Journal of Lightwave Technology 30(24), 3857–3879 (2012).
3. A. Carena et al., **“EGN model of non-linear fiber propagation”**, Optics Express 22, 16335–16362 (2014), DOI: 10.1364/OE.22.016335.

---

## Summary

이 폴더의 두 핵심 파일은 **범용 GN 기반 링크/NLI 및 시스템 성능 분석**을 위한 코드입니다. 변조 보정이 필요한 경우 제한적인 SCI-EGN 옵션을 사용할 수 있으며, **Full-EGN 연구 계산은 `EGN_model/EGN_adaptive.py`로 분리**되어 있습니다.

**Eng:** The two main files in this directory form a general GN-based NLI and system-performance toolkit. A limited SCI-EGN option is available, while the Full-EGN research implementation is maintained separately in `EGN_model/EGN_adaptive.py`.
