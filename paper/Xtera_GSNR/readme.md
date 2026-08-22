# Xtera Generalized SNR for Submarine Systems Reproduction

## White Paper

**Generalized SNR for Submarine Systems — Calculating performance in repeatered open line systems**

- Publisher: Xtera
- Type: White paper / technical presentation
- Source: [Xtera Knowledge Centre](https://xtera.com/knowledge-centre/)
- Scope: Generalized SNR prediction for repeatered submarine open line systems

이 디렉터리는 Xtera 백서의 **Fig. 2, Fig. 3, Fig. 4 및 Fig. 5**를 Python/Google Colab으로 재현합니다. 그래프 좌표를 수동 입력하거나 점별 보정하지 않고, 백서의 식 (3)–(7), 기준 채널 계획, 총 출력 제약과 GN 기반 비선형 모델로 주파수별 \(SNR_O\), \(SNR_{NL}\), \(SNR_G\)를 계산합니다.

## 논문 3줄 요약

1. 본 백서는 타사 coherent terminal을 수용하는 repeatered submarine open line system에서 케이블 성능을 특정 변조 형식이나 장비에 종속시키지 않고 **generalized SNR**로 표현하는 방법을 설명합니다.
2. 수신 품질을 ASE 중심의 선형 잡음 \(SNR_O\)와 Kerr 비선형 잡음 \(SNR_{NL}\)로 분리한 뒤 역수 합으로 \(SNR_G\)를 계산합니다.
3. 기준 channel plan에서 얻은 주파수별 total noise factor \(F_{total}(f)\)와 nonlinear coefficient \(K_{NL}(f)\)를 재사용하면 채널 수, 간격, baud rate 및 출력 tilt가 달라진 신규 terminal 구성의 성능을 빠르게 예측할 수 있습니다.

## Three-line Summary

1. The white paper describes a terminal-independent generalized-SNR framework for repeatered submarine open line systems.
2. Linear ASE noise and Kerr nonlinear noise are represented by \(SNR_O\) and \(SNR_{NL}\), then combined through reciprocal SNR addition.
3. Frequency-dependent \(F_{total}(f)\) and \(K_{NL}(f)\) extracted from a baseline link can be reused to predict different channel counts, spacings, baud rates, powers, and spectral tilts.

## 주요 백서 조건

| 항목 | 기준 조건 |
|---|---|
| Link 구성 | 25 spans × 140 km = 3,500 km |
| Span loss | 25.6 dB/span |
| 증폭 방식 | Raman repeatered link |
| 기준 채널 계획 | 100 channels × 50 GHz |
| 기준 symbol rate | 34 GBaud |
| 기준 점유 대역 | 약 4.95 THz |
| Repeater total output | 18 dBm |
| 기준 spectral tilt | 4 dB |
| Case A / Fig. 3 | 133 channels × 37.5 GHz, 34 GBaud, 4 dB tilt |
| Case B / Fig. 4 | 67 channels × 75 GHz, 34 GBaud, 2.5 dB tilt |
| Case C / Fig. 5 | 100 channels × 50 GHz, 50 GBaud, 4 dB tilt |

## 핵심 식과 물리적 의미

### 1. 누적 noise factor

동일 span이 \(N\)개 직렬 연결되고 각 span의 net gain이 loss를 보상할 때 백서의 누적 noise factor는 다음과 같이 계산됩니다.

\[
F_{total}(f)=N F_1(f)-(N-1)
\]

여기서 \(F_1(f)\)은 단일 span의 주파수별 noise factor입니다.

### 2. 선형 잡음 제한 SNR

\[
P_{ASE}(f)=F_{total}(f)h\nu B_{ch}
\]

\[
SNR_O(f)=\frac{P(f)}{P_{ASE}(f)}-\frac{1}{2}
\]

\(h\nu B_{ch}\)는 채널 대역 안에 들어오는 양자 잡음 항이며, \(-1/2\)는 장거리 constant-output amplifier chain에서의 signal-droop 보정입니다.

### 3. 비선형 잡음 제한 SNR

\[
\frac{1}{SNR_{NL}(f)}
=K_{NL}(f)\frac{P(f)^2}{B_{ch}^{2}}
\]

채널 출력 \(P\)가 커질수록 Kerr 비선형 penalty는 제곱으로 증가합니다. 반대로 동일 출력에서 baud rate \(B_{ch}\)가 커지면 power spectral density가 낮아져 비선형 제한 SNR은 개선됩니다.

### 4. Generalized SNR

\[
\frac{1}{SNR_G(f)}
=\frac{1}{SNR_O(f)}+\frac{1}{SNR_{NL}(f)}
\]

따라서 \(SNR_G\)는 선형 ASE와 비선형 잡음을 하나의 cable-performance 지표로 결합합니다. 실제 modem의 구현 penalty나 back-to-back SNR은 별도로 더해야 합니다.

### 5. 채널별 출력 분배

\[
P_{dBm}(f)
=P_{total,dBm}-10\log_{10}N_{ch}
+T_{dB}\frac{f-f_{mid}}{f_{max}-f_{min}}
\]

총 repeater 출력이 고정되므로 채널 수가 늘면 채널당 출력은 감소하고, 채널 수가 줄면 증가합니다. \(T_{dB}\)는 대역 전체의 launch-power tilt입니다.

## 핵심 결과

| 비교 | 백서와 알고리즘에서 확인되는 결과 |
|---|---|
| 기준 특성의 재사용 | 기준 plan에서 얻은 \(F_{total}(f)\)와 \(K_{NL}(f)\)를 새 channel grid에 보간하면 전체 link simulation을 매번 다시 수행하지 않고도 \(SNR_O\)와 \(SNR_G\)를 예측할 수 있습니다. |
| Case A: 100 → 133 channels | 고정 총출력에서 중심 채널 power가 약 1.25 dB 감소합니다. ASE 제한 SNR은 낮아지지만 채널당 power 감소로 nonlinear penalty는 완화됩니다. |
| Case B: 100 → 67 channels | 중심 채널 power가 약 1.72 dB 증가하여 \(SNR_O\)는 개선되지만, \(P^2\)에 비례하는 nonlinear penalty는 커집니다. 따라서 \(SNR_G\)는 두 효과의 절충으로 결정됩니다. |
| Case C: 34 → 50 GBaud | noise bandwidth 증가로 \(SNR_O\)는 약 1.67 dB 낮아지는 반면, 동일 \(K_{NL}\)과 출력에서는 \(B_{ch}^{2}\) 항 때문에 \(SNR_{NL}\)가 약 3.35 dB 개선됩니다. |
| Open-line 활용 | terminal 종류나 channel plan이 바뀌어도 공통 GSNR 기준으로 성능과 margin을 비교할 수 있어, 케이블 공급사와 terminal 공급사가 다른 환경의 upgrade 검토에 적합합니다. |

## 그림별 의미

| Figure | Channel plan | 주요 내용 |
|---|---|---|
| Fig. 2 | 100 × 50 GHz, 34 GBaud | 기준 link의 \(SNR_O\)와 \(SNR_G\)를 주파수별로 계산하고 link simulation과 비교 |
| Fig. 3 | 133 × 37.5 GHz, 34 GBaud | 채널 수 증가와 채널당 power 감소가 선형·비선형 SNR에 미치는 영향 |
| Fig. 4 | 67 × 75 GHz, 34 GBaud | 채널 수 감소와 채널당 power 증가에 따른 ASE 개선 및 nonlinear penalty의 trade-off |
| Fig. 5 | 100 × 50 GHz, 50 GBaud | baud rate 증가에 따른 ASE noise bandwidth와 nonlinear PSD 효과 비교 |

## Figure Reproduction Notebook

| Figures | Notebook | 설명 |
|---|---|---|
| Figs. 2–5 | [generalized_snr_fig2_fig5_reproduction.ipynb](./generalized_snr_fig2_fig5_reproduction.ipynb) · [Open in Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/paper/Xtera_GSNR/generalized_snr_fig2_fig5_reproduction.ipynb) | 백서 식 (3)–(7), 누적 noise factor, GN 기반 \(K_{NL}\), channel-power 배분 및 baseline 재사용 알고리즘으로 네 가지 channel plan을 계산합니다. |

## 재현 알고리즘

1. **Channel grid 생성:** 채널 수와 spacing으로 중심 주파수 기준의 WDM frequency grid를 생성합니다.
2. **출력 분배:** 총출력 18 dBm을 채널 수로 나누고 지정된 spectral tilt를 적용해 \(P(f)\)를 계산합니다.
3. **선형 잡음 계산:** 단일-span \(F_1(f)\)에서 \(F_{total}(f)\)를 구하고 \(P_{ASE}=F_{total}h\nu B_{ch}\)로 \(SNR_O\)를 계산합니다.
4. **비선형 계수 계산:** attenuation, dispersion, nonlinear coefficient, effective length 및 SCI/XCI 주파수 간격을 사용하는 closed-form incoherent GN surrogate로 \(K_{NL}(f)\)를 계산합니다.
5. **기준 특성 저장:** Fig. 2 baseline의 \(F_{total}(f)\)와 \(K_{NL}(f)\)를 기준 cable characterization으로 저장합니다.
6. **새 plan으로 보간:** Case A/B/C의 새 frequency grid에 기준 특성을 보간하고, 새 채널 출력과 baud rate를 식 (4)–(7)에 적용합니다.
7. **GSNR 결합:** \(SNR_O\)와 \(SNR_{NL}\)의 역수를 더해 \(SNR_G\)를 계산합니다.
8. **독립 검증:** baseline 특성을 재사용한 symbol 결과와 각 plan에서 GN kernel을 다시 계산한 simulation line을 비교합니다.

## 현재 노트북의 중심 채널 결과

아래 값은 백서의 원시 simulator 데이터가 아니라, 공개 조건과 대표적인 SMF-28e+ LL 파라미터를 사용한 **현재 Colab surrogate의 계산 결과**입니다.

| Case | \(SNR_O\) | \(SNR_G\), baseline \(K_{NL}\) 재사용 | \(SNR_G\), plan별 GN 재계산 |
|---|---:|---:|---:|
| Fig. 2 Baseline | 9.36 dB | 8.96 dB | 8.96 dB |
| Fig. 3 Case A | 8.03 dB | 7.86 dB | 7.86 dB |
| Fig. 4 Case B | 11.17 dB | 9.97 dB | 9.97 dB |
| Fig. 5 Case C | 7.56 dB | 7.44 dB | 7.44 dB |

## 재현 범위와 한계

- 백서의 실제 Raman pump power profile, span별 gain ripple, 정확한 \(A_{eff}\), \(\gamma\), \(\beta_2\), Fig. 1의 원시 \(F_1(f)\) 및 \(K_{NL}(f)\) 데이터는 공개되어 있지 않습니다.
- 따라서 현재 노트북은 25.6 dB/140 km 조건에 맞춘 smooth Raman noise-factor model과 대표 SMF 파라미터를 사용합니다.
- \(K_{NL}\)는 distributed-Raman power profile을 적분하는 공급사 simulator가 아니라 effective-length 기반 closed-form incoherent GN surrogate로 계산합니다.
- 논문 그림의 좌표를 digitizing하거나 점별 curve fitting 값으로 입력하지 않았습니다. 공개식이 설명하는 방향, 곡선 형상과 channel-plan trade-off를 독립적으로 재현하는 것이 목적입니다.
- 정밀 일치를 위해서는 노트북의 span noise-factor 함수와 nonlinear coefficient 함수를 commissioning 데이터 또는 공급사 GN/GGN 결과로 교체해야 합니다.

## Notes

- 노트북은 Google Colab에서 독립적으로 실행할 수 있습니다.
- NumPy와 Matplotlib만 사용하며 별도의 상용 simulator나 solver가 필요하지 않습니다.
- 코드의 계산 구조는 유지한 채 실제 cable acceptance 데이터의 \(F_{total}(f)\), \(K_{NL}(f)\), loss, dispersion 및 Raman gain profile로 교체할 수 있습니다.
