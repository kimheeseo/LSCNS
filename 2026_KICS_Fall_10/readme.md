# GN/EGN 적분 엔진과 변조 성능 검증

이 폴더는 WDM GN 적분 엔진, 변조/성능 레이어, 수식 검증 노트북, 그리고 첨부 논문 **A Simple and Accurate Closed-Form EGN Model Formula**의 Fig. 1–3 재현 보고서를 함께 제공한다.

## 파일별 특징

| 파일 | 역할과 특징 |
|---|---|
| gn_integral_general.py | 핵심 수치 엔진. 전체 WDM PSD에서 SCI/XCI/MCI를 자연스럽게 적분하고, 직사각/raised-cosine/RRC/custom PSD, β₂+β₃, 이기종 span, coherent/incoherent 누적, lumped EDFA·분포 이득, scrambled-Sobol QMC를 지원한다. 기본 경로는 순수 GN이며 egn_mu4/egn_mu6를 명시할 때만 제한적인 SCI EGN을 사용한다. |
| gn_integral_general_modulation.py | 엔진 어댑터와 시스템 성능 레이어. BPSK~256QAM의 명시적 성상도에서 μ₄/μ₆/Φ/Ψ를 직접 계산하고, ASE·NLI·GSNR·AWGN BER·gross/net rate·Shannon-gap 용량 및 launch-power sweep을 제공한다. NLI 적분은 엔진에 위임하며 기본값은 nli_model=\"gn\"이다. |
| GN_integral_math_verification.ipynb | GN 적분식의 단위·대칭성·극한·P³ power law·Sobol 수렴을 작은 독립 구현으로 검증한다. API 사용 예제가 아니라 수식/수치 sanity check용이다. |
| GN_integral_usage_guide.ipynb | 기존 GN API의 입력/출력, SMF·G.654.E 예제, span·launch-power sweep, self-test, 지원 범위와 한계를 단계별로 설명한다. |
| GN_modulation_Fig123_validation_colab.ipynb | 본 보고서. 첨부 PDF에서 벡터 곡선을 직접 판독한 기준값과 현재 modulation→engine 경로의 결과를 표·그래프로 비교한다. Colab에서 캐시를 먼저 볼 수 있고, 플래그를 바꾸면 동일 조건을 다시 계산할 수 있다. |
| paper_fig12_reference.json, paper_fig3_reference.json | PDF Fig. 1–3의 축 눈금과 벡터 선/마커에서 추출한 기준 배열. fitting으로 생성하지 않았으며, 파일 안에 페이지·추출 방법·판독 정밀도를 기록했다. |
| fig12_results_avg.json, fig3_code_gn_seeds.json | 고정 입력으로 미리 실행한 결과 캐시. 보고서의 기본 실행은 이 값을 표시하며, 재계산 플래그로 새 결과를 만들 수 있다. |

## 논문 조건과 비교 방법

- Fig. 1: 3 PM-QPSK, Fig. 2: 15 PM-QPSK, 32 GBd, 33.6 GHz spacing, raised-cosine roll-off 0.05, 100 km span, −3 dBm/channel, SMF·NZDSF·LS.
- Fig. 3: 15 channels, 32 GBd, 33.6/35/40/45/50 GHz spacing, roll-off 0.05, EDFA NF=5 dB. PM-QPSK는 120 km span와 BER 1.7×10⁻³, PM-16QAM은 85 km span와 BER 2×10⁻³를 사용한다. PSCF·SMF·NZDSF·LS의 논문 파라미터를 그대로 넣었다.
- Fig. 1–2 오차는 논문 GN 곡선 대비 |code−paper| [dB]와 선형 NLI 상대오차를 함께 계산한다. Fig. 3은 고정 −10…+5 dBm, 0.25 dB launch-power grid에서 최대 정수 passing span을 구하고, Sobol seed 11·22·33의 평균±표준편차를 보고한다.
- 난수 seed, grid, 적분점, fiber 파라미터를 결과를 본 뒤 조정하지 않았다. fitting 또는 curve-shape tuning은 없다.

## 결과 요약

### Fig. 1–2 (paper GN 곡선 대비)

| 구간 | 평균 절대오차 [dB] | 최대 절대오차 [dB] | 평균 상대오차 [%] | 최대 QMC 표준편차 [dB] |
|---|---:|---:|---:|---:|
| Fig. 1 전체 | 0.0281 | 0.0703 | 0.6474 | 0.0218 |
| Fig. 2 전체 | 0.0509 | 0.1562 | 1.1637 | 0.1639 |
| Fig. 1–2 전체 | **0.0395** | **0.1562** | **0.9052** | **0.1639** |

Fig. 2 SMF의 50-span 점에서 QMC seed 분산이 가장 크다. 보고서 그래프의 error bar는 이 수치 변동을 그대로 표시한다.

### Fig. 3 (paper GN 곡선 대비)

| 변조 | 평균 절대오차 [span] | 최대 절대오차 [span] | 평균 상대오차 | 최대 상대오차 | 최대 seed 표준편차 [span] |
|---|---:|---:|---:|---:|---:|
| QPSK | 0.8719 | 2.0009 | 3.7957% | 8.0839% | 2.8868 |
| 16QAM | 0.5686 | 1.5557 | 4.5149% | 16.3146% | 1.5275 |

논문 SIM 곡선과의 평균 상대 차이는 QPSK 14.5299%, 16QAM 13.2964%이다. 이는 현재 코드가 GN 적분+AWGN BER 근사이며 논문의 SSFM 시뮬레이터가 아니기 때문에 생기는 모델 차이로, 오차를 숨기거나 fitting하지 않았다.

## 해석상 주의점

현재 엔진의 EGN 옵션은 직사각 스펙트럼에서 CUT SCI만 보정하며 XCI/MCI EGN은 구현하지 않는다. 논문 Fig. 1–3은 roll-off 0.05와 full-WDM EGN/SIM을 사용하므로, 보고서의 주 비교 대상은 재현 가능한 **paper GN** 곡선이다. 논문의 EGN/SIM 선은 같은 그래프에 참고용으로 표시하되 현재 구현이 이를 재현했다고 주장하지 않는다. full-WDM EGN 또는 SSFM 검증에는 별도 구현이 필요하다.

또한 논문 곡선은 그래프 판독값(약 0.1 dB/0.1 span 정밀도)이고 코드 Fig. 3 결과는 정수 span pass/fail이다. 이 판독·양자화 한계를 결과 해석에 포함했다.

## 실행

GN_modulation_Fig123_validation_colab.ipynb를 Colab에서 열고 순서대로 실행한다. 기본값은 캐시 결과를 즉시 표시한다.

- RUN_FIG12=True: Fig. 1–2를 3개 Sobol seed로 재계산한다.
- RUN_FULL_FIG3=True: Fig. 3의 40개 조건을 고정 power grid로 재계산한다(Colab에서 수 분 소요).

입력/결과 JSON은 노트북과 같은 폴더에 있으므로 외부 경로 없이 재현할 수 있다. 사용한 첨부 PDF SHA-256은 e90abf9a9927cf82e7578b4e74f9b1d06a2956d0911c505921ceabe74a3ba40f이다.

## 참고 논문

1. P. Poggiolini et al., *A Simple and Accurate Closed-Form EGN Model Formula* (첨부 PDF, Fig. 1–3).
2. R. Dar et al., *Properties of nonlinear noise in long, dispersion-uncompensated fiber links*, Optics Express 21 (2013), https://doi.org/10.1364/OE.21.025685.
3. A. Carena et al., *EGN model of nonlinear fiber propagation*, Optics Express 22 (2014), https://doi.org/10.1364/OE.22.016335.
