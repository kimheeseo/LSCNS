# GN 모델 구현 완성도 평가

코드 신뢰성을 확인하기 위해 **Poggiolini 등의 「A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems」[1]과 두 Python 파일의 수식·가정을 비교**했습니다.

- [gn_integral_general.py](../gn_integral_general.py): WDM GN 적분과 링크 물리 계산.
- [gn_integral_general_modulation.py](../gn_integral_general_modulation.py): ASE·GSNR·BER 및 입사전력별 성능 계산.
- 평가일: **2026-10-01** / 논문 버전: **arXiv:1209.0394v13**.

## 1. 평가 결론

**핵심 적분형 GN 엔진의 주요 구조와 선정 수식의 재현성을 확인했습니다.** WDM PSD, 단일·다중 스팬 링크 함수, NLI의 전력 세제곱 관계와 ASE·GSNR·QPSK BER 연결을 확인했습니다.

논문의 모든 확장 조건이나 실험·SSFM 대비 절대 예측 정확도를 보증하는 결과는 아닙니다.

## 2. 논문과 코드의 대응

수식 번호는 [1]의 v13 기준입니다.

| 비교 항목 | 논문 근거 | 코드 위치 | 확인 범위 |
|---|---|---|---|
| 광전력·전계 감쇠 구분 | 제II절 | `alpha_field_from_db()` | 변환식 실행 |
| DP NLI 이중 적분, 16/27 계수 | 식 (88), (96) | `gn_nli_psd_qmc()` | QMC 표본 수렴 |
| beta2/beta3 위상 부정합 | 식 (G.1)–(G.4) | `phase_mismatch_beta23()` | 구현; beta2 실행 |
| 단일 스팬 비선형 소스 | 식 (88) | `_local_span_source_integral()` | 집중형 EDFA 경로 실행 |
| 다중 스팬 코히어런트·비코히어런트 누적 | 식 (96), (100), (101), 부록 B | `link_kernel()` | 동일 스팬 및 손실 보상된 서로 다른 2스팬 실행 |
| 분포형 증폭 | 식 (I.1), (103)–(105) | `distributed_field_log_gain()` | 프로파일 구현; 수치 검증 제외 |
| EDFA ASE와 SNR | 식 (8), (11) | `evaluate_performance()` | 표본 조건 실행 |
| 수신 대역 NLI 적분 | 식 (12) | `integrate_nli_over_channel()` | 폭 Rs의 직사각형 대역 |
| PM-QPSK BER | 식 (6) | `ber_awgn_approx()` | 표본 SNR 실행 |

## 3. 실행 검증 결과

### 핵심 GN 수식

두 대상 파일을 변경하지 않고 [paper1_reproducibility.py](./paper1_reproducibility.py)를 실행했습니다.

**7개 항목 모두 통과:** 감쇠 변환, beta2 위상 부정합, 단일 스팬 링크 함수, 동일 스팬의 비코히어런트 누적·코히어런트 위상합, 서로 다른 스팬의 위상 이력, 모든 채널의 전력 2배에 따른 NLI 8배 관계.

최대 상대오차는 **1.29452 × 10⁻¹¹ %**입니다. 이는 선정 수식·대수 관계의 재현 오차이며 실제 링크 예측 오차가 아닙니다. 일부 검증은 코드의 보조 함수를 공유하므로 완전히 독립적인 솔버와의 비교로 해석할 수 없습니다.

### Sobol 적분 수렴

조건: 3채널, 50 GHz 간격, 32 GBd, 채널당 −3 dBm, 80 km × 2스팬, 감쇠 0.2 dB/km, D=17 ps/(nm·km), gamma=1.3 /(W·km), 코히어런트 누적. 중앙 주파수 NLI PSD를 시드 3·11·29·47로 계산했습니다.

| 표본 수 | 시드 간 상대 표준편차 | 직전 표본 수 대비 평균값 변화 |
|---:|---:|---:|
| 2,048 | 2.2520% | — |
| 8,192 | 0.2223% | 2.2961% |
| 32,768 | 0.09349% | 0.14409% |

해당 조건·주파수에서의 수치 안정성을 확인했습니다. 표준편차는 절대오차 상한이 아니며, 다른 링크 조건과 수신 대역 적분의 수렴은 별도 검증이 필요합니다.

### 시스템 성능 모듈

3채널·32 GBd·80 km × 2스팬·NF 5 dB·TRX SNR 18 dB에서 추가 점검했습니다. NLI 계산은 Sobol 표본 4,096개, 시드 1, 수신 적분점 3개를 사용했습니다.

| 검증 항목 | 결과 |
|---|---:|
| EDFA ASE와 별도 계산한 식 (8)의 상대오차 | 4.44 × 10⁻¹⁶ |
| QPSK BER과 식 (6)의 최대 절대오차(SNR=1, 10, 100) | 0 |
| ASE+NLI+TRX로 재계산한 GSNR 차이 | 0 dB |
| 세제곱 scaling과 전력별 재적분 GSNR 차이(−3, 0 dBm) | 0 dB |

TRX는 추가 모델입니다. 이 점검 코드는 임시 실행했으며 저장소 재실행 스크립트에는 포함하지 않았습니다.

## 4. 적용 범위와 남은 검증

**전체 확장 조건과 실험·SSFM 대비 정확도 검증은 미완료입니다.**

- **분산·수신 필터:** beta2/beta3와 폭 Rs의 직사각형 수신 대역을 사용합니다. 일반 주파수 의존 전파상수 및 임의 수신 필터는 미구현입니다.
- **분포형 증폭:** GN 소스에 이득 프로파일을 적용할 수 있지만 ASE 함수는 집중형 EDFA용입니다. Raman 시스템 전체 GSNR은 검증하지 않았습니다.
- **계산 방식:** 제V절의 폐형식 근사 전체 대신 GN 적분을 직접 수치 계산합니다.
- **추가 기능:** Sobol QMC, TRX 잡음, 변조별 전송률, Shannon-gap 용량은 별도 평가 대상입니다. `nli_model="egn_sci"`는 SCI만 보정하며 전체 WDM EGN 검증을 뜻하지 않습니다.
- **운용 가정:** 중간 채널 추가·삭제, PMD/PDL, 상세 DSP는 포함하지 않습니다. 강한 비선형·분산 관리 링크로 결론을 일반화할 수 없습니다.

후속 검증 대상은 beta3, 불완전 손실 보상, 분포형 이득, 사용자 정의 PSD, 수신 적분점 수렴 및 동일 조건의 독립 적분·SSFM·실험 비교입니다.

## 참고문헌

[1] P. Poggiolini, G. Bosco, A. Carena, V. Curri, Y. Jiang, and F. Forghieri, **“A Detailed Analytical Derivation of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems,”** arXiv:1209.0394, v13, 2014. [논문](https://arxiv.org/abs/1209.0394v13) · [PDF](https://arxiv.org/pdf/1209.0394v13).
