# HCF Horseshoe-and-Spur Network Reproduction

## Paper

**Enhanced Scalability of Horseshoe-and-Spur Networks by Exploiting Hollow-Core Fiber**

- Authors: Mohammad M. Hosseini, João Pedro, Antonio Napoli
- Preprint: arXiv:2608.07082v1, 7 August 2026
- Source: [arXiv abstract](https://arxiv.org/abs/2608.07082) · [HTML](https://arxiv.org/html/2608.07082v1) · [PDF](https://arxiv.org/pdf/2608.07082)

이 디렉터리는 논문의 **Fig. 2, Fig. 3, Fig. 4 및 Fig. 5**를 Python/Google Colab으로 재현합니다. 논문 그래프의 좌표를 계산 입력이나 보정값으로 사용하지 않고, 논문에 공개된 식·전송 조건·최적화 제약과 명시적인 네트워크 표본 생성 알고리즘으로 결과를 계산합니다.

## 논문 3줄 요약

1. 본 논문은 DSCM 기반 coherent point-to-multipoint 송수신기를 사용하는 filterless horseshoe-and-spur 메트로·액세스망에서 기존 solid-core fiber(SCF)를 hollow-core fiber(HCF)로 전환했을 때의 전력 예산과 확장성을 분석합니다.
2. HCF의 매우 낮은 비선형성은 설계 병목을 SCF의 Kerr 비선형 임계값에서 광증폭기의 총 출력 한계로 이동시키며, SCF 대비 spur 전력 예산을 최대 약 20 dB 높입니다.
3. 5개 transit node 조건에서는 증폭기 수가 약 11–13개를 넘으면 이득이 포화되고, main horseshoe는 SCF로 유지하면서 spur만 HCF로 바꾸는 hybrid 구조도 증폭기 절감으로 HCF 비용 프리미엄을 상쇄할 수 있음을 보입니다.

## Three-line Summary

1. The paper evaluates filterless horseshoe-and-spur metro-access networks that combine coherent DSCM point-to-multipoint transceivers with full or targeted HCF deployment.
2. HCF moves the dominant power constraint from the SCF Kerr-nonlinearity threshold to aggregate amplifier output power and delivers up to approximately 20 dB more spur power budget than SCF.
3. Amplifier benefits saturate at roughly 11–13 units in a five-transit-node network, while a hybrid SCF horseshoe with HCF spurs can recover the HCF cost premium through amplifier savings.

## 주요 논문 조건

| 항목 | 논문 조건 |
|---|---|
| Transit node 수 | 5 |
| 400G DSCM 채널 수 | 10, 20 |
| 채널당 subcarrier 수 | 16 |
| 변조 형식 | 16QAM |
| EDFA 최대 총 출력 | 27 dBm, 37 dBm |
| 송신 출력 | −12 dBm/SC |
| 수신 감도 | −24 dBm/SC |
| SCF 비선형 전력 임계값 | −8 dBm/SC |
| HCF 비선형 전력 임계값 | +10 dBm/SC |
| 허용 SC 전력 불균형 | 8 dB |
| SCF/HCF 감쇠 | 각각 0.24 dB/km |
| HCF splice loss | 0.2 dB |
| Coupler excess loss | 0.5 dB |
| 평균 horseshoe span | 12 km |
| Coupler 후보 | Balanced 50:50 또는 unbalanced 50:50–90:10, 10% 간격 |
| 네트워크 평가 표본 | 10개 |
| 신뢰구간 | 평균 결과의 90% confidence interval |

## 핵심 결과

| 비교 항목 | 논문의 결론 |
|---|---|
| Full HCF 대 SCF | HCF가 spur당 전력 예산을 최대 약 20 dB 개선 |
| Balanced 대 unbalanced coupler | Unbalanced coupler 최적화가 더 높은 전력 예산 제공 |
| 증폭기 밀도 | 약 11–13개 이후 추가 증폭기의 한계효용이 크게 감소 |
| Hybrid SCF–HCF | SCF main horseshoe와 HCF spur의 성능은 SCF와 full HCF 사이이며, 증폭기가 10개를 넘으면 full HCF와의 차이가 거의 사라짐 |
| 경제성 예시 | HCF spur 300 km, HCF 프리미엄 50 USD/km, 증폭기 TCO 5,000 USD일 때 증폭기 3개 절감으로 15,000 USD 프리미엄 상쇄 |

## 그림별 의미

| Figure | 논문에서의 성격 | 주요 내용 |
|---|---|---|
| Fig. 2 | 해석식 기반 출력 한계 | DSCM 채널 수와 증폭기 최대 총 출력에 따른 subcarrier당 출력 상한 |
| Fig. 3 | ILP 네트워크 최적화 | 증폭기 수에 따른 평균 spur 전력 예산을 SCF·HCF, balanced·unbalanced coupler, uniform·nonuniform budget 조건별 비교 |
| Fig. 4 | ILP 네트워크 최적화 | Unbalanced coupler에서 full HCF와 hybrid SCF–HCF의 전력 예산 차이 |
| Fig. 5 | 해석식 기반 passive spur 모델 | Single-stage와 multi-stage tree의 노드 수·거리·필요 전력 예산 trade-off |

## Figure Reproduction Notebook

| Figure | Notebook | 1줄 설명 |
|---|---|---|
| Figs. 2, 3, 4, 5 | [horseshoe_fig2_fig3_fig4_fig5_reproduction.ipynb](./horseshoe_fig2_fig3_fig4_fig5_reproduction.ipynb) · [Open in Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/paper/Horseshoe/horseshoe_fig2_fig3_fig4_fig5_reproduction.ipynb) | 공개 수식과 제약으로 Fig. 2·5를 직접 계산하고, 재현 가능한 MILP 대리모델로 Fig. 3·4의 물리적 추세와 포화 현상을 계산합니다. |

## 재현 알고리즘

- **Fig. 2:** 400G DSCM 채널 하나가 16개 SC를 사용한다는 총출력 제약 `16 × N_Ch × P_SC,out ≤ P_A`를 dBm 영역의 `P_SC,out = P_A − 10 log10(16N_Ch)`로 변환해 모든 격자점을 직접 계산합니다.
- **Figs. 3·4:** 5개 transit node의 양방향 horseshoe에서 main-link 및 spur amplifier의 배치·이득, discrete coupler ratio, 수신감도, 최대 전력 불균형, amplifier 총출력 및 fiber 비선형 상한을 혼합정수선형계획법(MILP)으로 동시에 최적화합니다.
- **Uniform budget:** 모든 spur가 동일한 전력 예산을 갖도록 equality constraint를 적용합니다.
- **Nonuniform budget:** spur별 예산을 다르게 허용하면서 평균 전력 예산을 최대화합니다.
- **네트워크 표본:** 공개된 평균 12 km와 분포 통계를 따르는 log-normal surrogate에서 층화 표본추출하고 고정 random seed를 사용하므로 실행할 때마다 동일한 10개 네트워크가 생성됩니다.
- **Full HCF / Hybrid:** Full HCF는 main horseshoe와 spur를 모두 HCF로, hybrid는 main horseshoe를 SCF로 유지하고 spur만 HCF로 구성합니다.
- **Fig. 5:** `a = 10^(−αL/10)`와 논문의 식 (1), (2)로부터 `P_b,single = αL + 10log10(N)` 및 `P_b,multi = NαL + 10(N−1)log10(2)`를 계산합니다.

## 재현 범위와 한계

- Fig. 2와 Fig. 5는 논문에 식이 완전하게 공개되어 있어 직접 재현합니다.
- Fig. 3와 Fig. 4에 사용된 저자들의 전체 ILP 구현, 선행 연구의 실제 10개 네트워크 링크 길이, 각 최적해의 원시 데이터는 공개되어 있지 않습니다.
- 따라서 Fig. 3와 Fig. 4는 논문의 공개 제약을 반영한 독립적인 MILP surrogate입니다. SCF/HCF의 상대적 차이, coupler 최적화 효과, 증폭기 수에 따른 증가와 포화 같은 경향은 재현하지만 각 데이터 점은 원문과 다를 수 있습니다.
- 논문 그림에서 수동으로 읽은 좌표, 임의의 curve-fitting 목표값 또는 결과를 맞추기 위한 점별 보정값은 사용하지 않습니다.
- 논문과 동일하게 SCF와 HCF의 감쇠를 0.24 dB/km로 두어 HCF의 이점이 낮은 Kerr 비선형성에서 오도록 비교합니다.
- 논문은 대상 short-reach 환경에서 누적 ASE가 충분히 낮다고 전제합니다. 이 노트북도 상세 EDFA ASE 스펙트럼이나 광전송 파형 시뮬레이션 대신 전력 흐름과 배치 최적화에 초점을 둡니다.

## Notes

- 노트북은 Google Colab에서 독립적으로 실행할 수 있습니다.
- 기본값은 논문과 같은 10개 네트워크 ensemble이며, 계산 시간 단축이 필요한 smoke test에만 환경변수로 표본 수를 줄일 수 있습니다.
- SciPy의 MILP solver를 사용하므로 상용 최적화 solver나 별도 라이선스가 필요하지 않습니다.
