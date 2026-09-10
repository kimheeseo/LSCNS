# Horseshoe-and-Spur HCF Network Reproduction

## Paper

**Enhanced Scalability of Horseshoe-and-Spur Networks by Exploiting Hollow-Core Fiber**

- Authors: Mohammad M. Hosseini, João Pedro, Antonio Napoli
- Affiliations: Nokia Optical Networks; Instituto de Telecomunicações / Instituto Superior Técnico
- Source: [arXiv:2608.07082](https://arxiv.org/abs/2608.07082)
- Submitted: 7 August 2026

이 디렉터리는 논문의 **Fig. 2, Fig. 3, Fig. 4 및 Fig. 5**를 Python/Google Colab으로 계산하고 원 논문 결과와 비교합니다. 재현 곡선은 논문 그래프 좌표를 입력하거나 역보정하지 않고, 공개된 식과 Table I 조건 및 명시적인 대체 topology를 사용해 생성합니다.

## 논문 3줄 요약

1. 본 논문은 HCF의 매우 낮은 비선형성을 활용하여 DSCM 기반 filterless horseshoe-and-spur 메트로·액세스망의 EDFA 배치와 spur 전력예산을 최적화합니다.
2. HCF는 시스템 제한을 SCF의 Kerr 비선형성에서 EDFA 총 출력으로 이동시켜 spur 전력예산을 최대 약 20 dB 높이고, 증폭기 11–13대 부근에서 증가 효과가 포화됨을 제시합니다.
3. 선택적 HCF 적용은 필요한 증폭기 수를 줄여 HCF의 초기 가격 프리미엄을 상쇄할 수 있으며, 300 km에 50 USD/km의 추가비용을 가정하면 증폭기 3대 절감으로 손익분기에 도달합니다.

## Colab Notebook

| Reproduction | Notebook | Method |
|---|---|---|
| Figs. 2–5 | [horseshoe_fig2_5_reproduction_colab.ipynb](./horseshoe_fig2_5_reproduction_colab.ipynb) · [Open in Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/paper/horseshoe/horseshoe_fig2_5_reproduction_colab.ipynb) | DSCM 총 출력 수식, EDFA 배치·이득 MILP, 커플러 손실 및 spur tree 전력예산 식으로 계산하고 arXiv SVG의 논문 마커와 자동 비교합니다. |

## Figure Reproduction Method

| Figure | Calculation |
|---|---|
| Fig. 2 | `P_SC,out = P_A − 10 log10(16 N_Ch)`을 2차원 grid에 적용하여 논문과 같은 heatmap 및 네 조건의 계산값을 생성합니다. |
| Fig. 3 | 22개 후보 EDFA 위치와 continuous gain, 5개 transit node, 수신감도, SCF/HCF 출력 임계값, fiber·coupler·splice loss 및 8 dB imbalance를 제약으로 하는 SciPy MILP로 평균 spur budget과 90% CI를 계산합니다. Unbalanced ratio도 50:50–90:10 집합에서 one-hot 변수로 최적화합니다. |
| Fig. 4 | 동일 MILP에서 full-HCF와 SCF-ring/HCF-spur hybrid를 각각 풀어 `hybrid − full HCF` 전력예산 차이를 계산합니다. |
| Fig. 5 | 논문 식 (1), (2)에 `a=10^(−αL/10)`을 대입하여 single-stage와 cascaded 50:50 multi-stage tree의 node·거리별 요구 전력예산을 직접 계산합니다. |

## Model Scope and Limitations

- 이 논문은 **GN model을 사용하지 않습니다**. SCF와 HCF의 비선형 동작은 Table I의 −8 dBm/SC 및 10 dBm/SC hard limit로 모델링됩니다.
- 원 저자의 10개 topology별 link length와 전체 Julia/JuMP 제약식은 공개되지 않았습니다. 따라서 Fig. 3–4는 평균 link length가 정확히 12 km인 고정-seed 10개 log-normal topology ensemble을 사용한 독립 재현입니다.
- balanced coupler는 50:50, unbalanced case는 논문의 50:50–90:10 허용 범위에서 node·방향별 ratio를 MILP가 선택합니다. paper curve 좌표는 ratio나 EDFA 배치를 정하는 입력으로 사용하지 않습니다.
- 논문은 EDFA 총 출력만 제시하고 장비별 gain 범위는 공개하지 않습니다. 노트북은 비음수 gain과 느슨한 30 dB 수치 상한을 사용하며, per-subcarrier 출력 cap을 실제 물리 상한으로 적용합니다.
- Colab은 모델 계산을 먼저 완료한 후 arXiv HTML의 vector marker를 자동 추출하여 빈 사각형으로 겹쳐 그리고 MAE를 출력합니다. paper marker는 최적화 입력으로 사용되지 않습니다.
- GN/SSFM 확장을 위한 함수와 연결 지점은 포함하지만, 논문에 `γ`, `β₂`, baud rate, channel spacing, ASE/NLI 계수가 없으므로 임의값으로 Figures 2–5를 보정하지 않습니다.

## Main Paper Parameters

| Parameter | Value |
|---|---:|
| Transit nodes | 5 |
| DSCM channels | 10, 20 |
| Subcarriers per 400G channel | 16 |
| EDFA total output | 27, 37 dBm |
| Launch / receiver sensitivity | −12 / −24 dBm per SC |
| SCF / HCF power limit | −8 / 10 dBm per SC |
| Fiber attenuation | 0.24 dB/km |
| HCF splice loss | 0.2 dB |
| Coupler excess loss | 0.5 dB |
| Maximum SC imbalance | 8 dB |
| Mean horseshoe link length | 12 km |

## Notes

- 노트북은 Google Colab에서 위에서 아래로 독립 실행할 수 있으며 SciPy의 오픈소스 HiGHS MILP solver를 사용합니다.
- 기본값 `FULL_RUN=False`는 빠른 확인용 고정-seed topology 3개를 사용합니다. `FULL_RUN=True`로 바꾸면 논문과 같은 10개 ensemble 반복을 수행하며 Colab에서 더 오래 걸립니다.
- 기본 실행의 paper SVG marker 대비 MAE는 Fig. 3 약 3.26 dB, Fig. 4 약 5.58 dB입니다. 이는 공개되지 않은 원 topology·전체 ILP 제약으로 인한 차이를 숨기지 않고 정량화한 값이며, Fig. 2와 Fig. 5는 공개 식의 직접 계산입니다.
- 결과 CSV는 Colab 세션에 `fig3_milp_reproduction.csv`, `fig4_hybrid_difference_reproduction.csv`, `fig3_text_anchor_validation.csv`로 저장됩니다.
