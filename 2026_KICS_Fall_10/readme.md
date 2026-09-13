# 2026_KICS_Fall_10 — General GN Integral Model

이 폴더는 코히어런트 광전송 시스템의 비선형 간섭(NLI)을 수치 적분 기반 Gaussian Noise (GN) Model로 계산하기 위한 범용 코드와, 그 결과를 변조 방식·ASE·GSNR·BER·전송용량 평가까지 연결하는 성능 레이어를 포함합니다.

**English:** This directory contains a general-purpose numerical Gaussian Noise (GN) Model engine for nonlinear-interference (NLI) estimation in coherent optical transmission systems, together with a modulation/system-performance layer that extends the NLI result to ASE, GSNR, BER, and capacity analysis.

---

## 1. 파일 구성 / Files

### `gn_integral_general.py`

전체 WDM PSD를 직접 2차원 주파수 적분하여 NLI PSD와 채널 내 NLI 전력을 계산하는 핵심 수치 GN 엔진입니다. 단일 채널 전용 코드가 아니라 다채널 WDM 링크를 범용적으로 다룰 수 있도록 작성되었습니다.

**English:** This is the core numerical GN engine. It directly performs the two-dimensional frequency-domain GN integral over the full WDM power spectral density and computes NLI PSD or NLI power in a selected channel. It is designed as a reusable WDM/link engine rather than a single-channel-only implementation.

주요 기능은 다음과 같습니다.

**English:** Main capabilities include:

- 전체 WDM PSD 적분을 통해 SCI, XCI, MCI가 별도의 경험식 없이 적분 영역에서 자연스럽게 포함됩니다.  
  **English:** SCI, XCI, and MCI arise naturally from integration of the complete WDM PSD rather than from separate fitted formulas.
- 채널별 서로 다른 전력, 대역폭 및 비균일 채널 간격을 지원합니다.  
  **English:** Unequal channel powers, bandwidths, and irregular channel spacing are supported.
- rectangular, raised-cosine, RRC-power 및 사용자 정의 PSD를 지원합니다.  
  **English:** Rectangular, raised-cosine, RRC-power, and user-defined PSD shapes are supported.
- `beta2`뿐 아니라 `beta3`까지 포함한 위상 불일치를 계산할 수 있습니다.  
  **English:** Phase mismatch can include both `beta2` and `beta3`.
- span별 길이, 손실, 분산, 비선형계수, 증폭기 이득이 서로 다른 이기종 링크를 지원합니다.  
  **English:** Heterogeneous links with span-dependent length, loss, dispersion, nonlinear coefficient, and lumped gain are supported.
- coherent 또는 incoherent span accumulation을 선택할 수 있습니다.  
  **English:** Coherent or incoherent span accumulation can be selected.
- lumped EDFA뿐 아니라 ideal distributed gain, 단순 backward-Raman profile, custom gain profile을 지원합니다.  
  **English:** In addition to lumped EDFAs, ideal distributed gain, a simplified backward-Raman profile, and custom gain profiles are supported.
- 2-D 주파수 적분에는 scrambled Sobol quasi-Monte-Carlo(QMC)를 사용합니다.  
  **English:** Scrambled Sobol quasi-Monte-Carlo (QMC) sampling is used for the two-dimensional frequency integral.

### 입력과 출력 / Inputs and outputs

대표 입력은 WDM 채널의 중심주파수, launch power, baud rate, pulse shape와 각 span의 길이, 감쇠, `gamma`, `D` 또는 `beta2`, `beta3`, 증폭 조건입니다. 채널 전력과 PSD는 total dual-polarization 기준이며 GN coefficient는 코드에서 `16/27`을 사용합니다.

**English:** Typical inputs include WDM center frequencies, launch powers, baud rates, pulse shapes, and per-span length, attenuation, `gamma`, `D` or `beta2`, `beta3`, and amplification parameters. Channel power and PSD follow a total dual-polarization convention, and the code uses the DP GN coefficient `16/27`.

주요 출력은 특정 주파수의 NLI PSD와 선택 채널 대역 내 적분 NLI power입니다. 이후 `gn_integral_general_modulation.py`가 이 값을 시스템 성능 지표로 확장합니다.

**English:** Main outputs are NLI PSD at a target frequency and integrated NLI power over a selected channel. `gn_integral_general_modulation.py` then converts these results into system-level performance metrics.

### GN과 EGN 범위 / GN and EGN scope

기본 경로는 Gaussian-signal assumption을 사용하는 순수 GN Model입니다. 선택적으로 `egn_mu4`와 `egn_mu6`를 지정하면 Carena et al. (2014)의 직사각 스펙트럼 SCI 보정 항을 사용할 수 있지만, 이 파일만으로는 full-WDM EGN의 XCI/MCI 보정까지 구현된 것은 아닙니다.

**English:** The default path is the conventional GN Model under the Gaussian-signal assumption. Optional `egn_mu4` and `egn_mu6` parameters enable the rectangular-spectrum SCI correction of Carena et al. (2014), but this file alone does not implement full-WDM EGN corrections for XCI and MCI.

---

## 2. `gn_integral_general_modulation.py`

이 파일은 `gn_integral_general.py`의 NLI 엔진 위에 변조 방식과 시스템 성능 계산을 추가하는 상위 레이어입니다. 즉, NLI 적분 알고리즘 자체를 별도로 다시 구현하기보다는 핵심 적분 계산을 GN 엔진에 위임하고, 그 결과를 실제 통신 성능 지표로 변환합니다.

**English:** This file is a higher-level modulation and system-performance layer built on top of `gn_integral_general.py`. It does not duplicate the core NLI integration; instead, it delegates NLI calculation to the GN engine and converts the result into practical communication-system metrics.

지원 변조 방식은 BPSK, QPSK, 8QAM, 16QAM, 32QAM, 64QAM, 256QAM입니다.

**English:** Supported modulation formats are BPSK, QPSK, 8QAM, 16QAM, 32QAM, 64QAM, and 256QAM.

각 성상도에서 정규화된 `mu4`, `mu6`, `Phi = mu4 - 2`, `Psi = mu6 - 9mu4 + 12`를 직접 계산합니다. 따라서 변조 보정 계수를 임의 fitting 값으로 두지 않고 명시적으로 정의된 constellation에서 계산할 수 있습니다. 단, 8QAM과 32QAM은 코드에 명시된 특정 cross-constellation을 사용하므로 다른 성상도를 사용할 경우 moments를 다시 계산해야 합니다.

**English:** Normalized `mu4`, `mu6`, `Phi = mu4 - 2`, and `Psi = mu6 - 9mu4 + 12` are calculated directly from explicitly defined constellations rather than from fitted correction factors. The built-in 8QAM and 32QAM definitions are specific cross-constellations; moments must be recalculated if a different geometry is intended.

### 주요 성능 계산 / Main performance calculations

- EDFA ASE noise accumulation  
  **English:** EDFA ASE-noise accumulation
- 선택 CUT의 NLI power  
  **English:** NLI power in the selected channel under test (CUT)
- `SNR_ASE`, `SNR_NLI`, transceiver SNR를 포함한 GSNR  
  **English:** `SNR_ASE`, `SNR_NLI`, and GSNR including optional transceiver SNR
- modulation별 approximate AWGN BER  
  **English:** Approximate AWGN BER for each modulation format
- launch-power versus GSNR sweep  
  **English:** Launch-power-versus-GSNR sweep
- dual-polarization gross/net line rate  
  **English:** Dual-polarization gross and net line rates
- Shannon-gap 기반 capacity estimate  
  **English:** Shannon-gap-based capacity estimate

`nli_model="gn"`이 기본값이며 순수 GN 동작을 유지합니다. `nli_model="egn_sci"`를 선택하면 `mu4`, `mu6`를 핵심 엔진으로 전달하여 modulation-dependent SCI EGN correction을 계산합니다. 현재 이 경로에서도 XCI/MCI는 GN으로 유지됩니다.

**English:** `nli_model="gn"` is the default and preserves the pure-GN behavior. With `nli_model="egn_sci"`, `mu4` and `mu6` are passed to the core engine for modulation-dependent SCI EGN correction. XCI and MCI remain GN in this path.

---

## 3. 두 코드의 관계 / Relationship between the two files

```text
gn_integral_general.py
    └─ WDM PSD + link physics
       └─ GN frequency integral
          └─ NLI PSD / NLI power
                ↓
gn_integral_general_modulation.py
    └─ modulation moments
    └─ ASE
    └─ GSNR / BER
    └─ rate / capacity
    └─ launch-power sweep
```

`gn_integral_general.py`가 물리 계층의 핵심 NLI 계산기라면, `gn_integral_general_modulation.py`는 그 결과를 시스템 설계와 변조 비교에 사용할 수 있도록 확장한 코드입니다.

**English:** `gn_integral_general.py` is the physical-layer NLI engine, while `gn_integral_general_modulation.py` extends its output for system design, modulation comparison, and performance estimation.

---

## 4. 성능 및 검증 / Performance and validation

저장소의 `EGN_model/EGNvsGN/GN_modulation_Fig123_validation_colab.ipynb` 및 관련 JSON 결과에서는 동일 GN 적분 경로를 논문의 Fig. 1–3 조건과 비교했습니다. 저장된 검증 결과 기준으로 Fig. 1–2의 paper-GN 곡선 대비 전체 평균 절대오차는 약 **0.0395 dB**, 평균 선형 NLI 상대오차는 약 **0.9052%**로 보고되었습니다. 이는 해당 검증 조건에서 GN baseline이 논문 GN 결과를 높은 일관성으로 재현한다는 의미입니다.

**English:** The repository validation notebook `EGN_model/EGNvsGN/GN_modulation_Fig123_validation_colab.ipynb` compares the same GN-integral path with the paper conditions of Figs. 1–3. For Figs. 1–2, the stored validation results report an overall mean absolute error of approximately **0.0395 dB** against the paper GN curves and an average linear-NLI relative error of approximately **0.9052%**. This indicates strong consistency with the paper GN baseline under those tested conditions.

Fig. 3의 maximum-reach 비교에서는 QPSK 평균 절대오차 약 0.872 span, 16QAM 약 0.569 span이 보고되었습니다. 논문의 SSFM/SIM 결과와 차이가 더 큰 이유는 현재 코드가 GN 적분 + AWGN BER 근사를 사용하는 반면 논문 SIM은 split-step simulation 기반이기 때문입니다.

**English:** In the Fig. 3 maximum-reach comparison, the stored results report mean absolute errors of approximately 0.872 spans for QPSK and 0.569 spans for 16QAM against the paper GN curves. Larger differences versus the paper SSFM/SIM curves are expected because this code uses a GN integral plus an AWGN BER approximation rather than a split-step propagation simulator.

### 장점 / Strengths

- closed-form 근사식보다 계산 비용은 크지만, 전체 PSD와 실제 링크 구성을 직접 적분하므로 비균일 WDM과 이기종 span 등으로 확장하기 쉽습니다.  
  **English:** It is computationally heavier than a closed-form approximation, but direct PSD/link integration makes it easier to extend to irregular WDM grids and heterogeneous spans.
- 논문 결과에 맞추기 위한 curve fitting 없이 물리 파라미터를 입력해 계산하도록 구성되어 있습니다.  
  **English:** The implementation is parameter-driven and does not require curve fitting to paper targets.
- Sobol QMC seed와 적분 해상도를 변경하여 numerical convergence를 확인할 수 있습니다.  
  **English:** Numerical convergence can be checked by varying Sobol QMC seeds and integration resolution.

### 한계 / Limitations

- GN Model 자체가 Gaussian-signal assumption을 사용하므로 실제 변조 신호의 비가우시안 특성을 완전히 반영하지 않습니다.  
  **English:** The GN Model relies on the Gaussian-signal assumption and therefore does not fully capture the non-Gaussian statistics of actual modulation formats.
- 선택적 EGN 기능은 현재 SCI 중심이며 full-WDM XCI/MCI EGN이 아닙니다.  
  **English:** The optional EGN path is SCI-focused and is not a full-WDM XCI/MCI EGN implementation.
- distributed Raman 처리와 custom profile은 범용 수치 확장 기능이므로 실제 시스템 설계 sign-off 전에는 전용 Raman 모델 또는 상용 시뮬레이터와의 추가 검증이 필요합니다.  
  **English:** Distributed-Raman and custom-profile support are numerical extensions and should be independently validated against a dedicated Raman model or commercial simulator before design sign-off.
- QMC 기반 적분은 closed-form보다 계산 시간이 길고, 높은 정확도가 필요한 경우 seed 및 sample 수에 대한 convergence 확인이 필요합니다.  
  **English:** QMC integration is slower than closed-form evaluation and requires seed/sample convergence checks when high numerical accuracy is required.

---

## 5. 참고 논문 / References

### GN Model 기준 / GN Model basis

P. Poggiolini et al., **“A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems”**, arXiv:1209.0394.

이 코드는 해당 계열의 GN 적분 구조, span coherent accumulation, dispersion/phase-mismatch 및 distributed-gain formulation을 기반으로 일반화되어 있습니다.

**English:** The code follows the Poggiolini GN-model framework for numerical NLI integration, coherent span accumulation, dispersion/phase mismatch, and distributed-gain treatment, generalized for reusable WDM/link simulations.

### EGN / modulation correction

A. Carena, G. Bosco, V. Curri, Y. Jiang, P. Poggiolini, F. Forghieri, **“EGN model of non-linear fiber propagation”**, Optics Express 22, 16335–16362 (2014), DOI: 10.1364/OE.22.016335.

`egn_mu4`, `egn_mu6`, `Phi`, `Psi` 및 선택적 SCI-EGN correction의 직접적인 기준입니다.

**English:** This paper is the direct reference for the `mu4`, `mu6`, `Phi`, `Psi` definitions and the optional SCI-EGN correction implemented in the code.

R. Dar et al., **“Properties of nonlinear noise in long, dispersion-uncompensated fiber links”**, Optics Express 21 (2013), DOI: 10.1364/OE.21.025685.

비가우시안 modulation-dependent nonlinear correction의 물리적 해석과 normalized fourth-moment correction을 교차 확인하는 참고 문헌입니다.

**English:** This work is used as an additional reference for the physical interpretation of modulation-dependent nonlinear noise and normalized fourth-moment correction terms.

### Fig. 1–3 검증에 사용된 문헌 / Reference used for Fig. 1–3 validation

P. Poggiolini et al., **“A Simple and Accurate Closed-Form EGN Model Formula”**.

저장소의 `EGN_model/EGNvsGN/` 검증 노트북은 이 문헌의 Fig. 1–3에서 GN 곡선을 중심으로 현재 적분 코드와 비교하며, EGN 및 SIM 곡선은 구현 범위를 구분하기 위한 참고값으로 사용합니다.

**English:** The validation notebook under `EGN_model/EGNvsGN/` compares the current numerical integral mainly against the paper's GN curves in Figs. 1–3; EGN and SIM curves are retained as references to distinguish the current implementation scope.

---

## 6. 사용 시 주의사항 / Usage notes

1. `alpha_db_per_km`는 power attenuation [dB/km]이며 내부에서는 field attenuation으로 변환됩니다.  
   **English:** `alpha_db_per_km` is a power-attenuation value in dB/km and is internally converted to a field-attenuation coefficient.
2. 채널 launch power는 total dual-polarization power 기준입니다.  
   **English:** Channel launch power uses a total dual-polarization power convention.
3. `MODULATION_PHI`의 `Phi=mu4-2`와 `Channel.modulation_phi`는 의미가 다릅니다. 후자는 legacy SCI multiplier이며 기본값 1입니다.  
   **English:** `MODULATION_PHI` (`Phi=mu4-2`) is not the same quantity as `Channel.modulation_phi`; the latter is a legacy SCI multiplier whose default is 1.
4. 연구·비교용으로는 GN baseline을 우선 사용하고, EGN이 필요한 경우 `EGN_model/`의 별도 검증 결과와 제한사항을 함께 확인하십시오.  
   **English:** For general research and comparison, use the validated GN baseline first. If EGN behavior is required, also review the dedicated validation results and limitations under `EGN_model/`.
5. 새로운 광섬유 또는 저분산 링크에 적용할 경우 독립 SSFM/VPIphotonics 등의 교차 검증을 권장합니다.  
   **English:** For new fiber types or low-dispersion links, independent cross-validation with SSFM, VPIphotonics, or another trusted simulator is recommended.
