# Paper Reproduction

이 폴더는 논문에 제시된 수치, 알고리즘 및 그래프를 Python/Google Colab으로 재검증한 프로젝트를 모아두기 위한 공간입니다.

## Reproduction Projects

| Folder | Based on | Source |
|---|---|---|
| [hcf-optimum-launch-power-reproduction](./hcf-optimum-launch-power-reproduction/) | **On the Optimum Energy-per-bit Launch Power in Coherent Hollow-core Fibre Transmission Systems** | [arXiv:2606.17942](https://arxiv.org/abs/2606.17942) |
| [opticalfiber](./opticalfiber/) | **High Density Optical Cable with Ultra-Low-Loss, Large-Effective-Area ITU-T G.654.E Optical Fiber** | [IWCS Webinar 85](https://iwcs.org/webinar/85/) |
| [EDFA-Raman-320-km](./EDFA-Raman-320-km/) | **320㎞ Optical Transmission using EDFA and Raman amplifier for 10Gbit/s 128 Channel DWDM Signals** | [KICS Digital Library](https://journal.kics.or.kr/digital-library/33081) |
| [Horseshoe](./Horseshoe/) | **Enhanced Scalability of Horseshoe-and-Spur Networks by Exploiting Hollow-Core Fiber** | [arXiv:2608.07082](https://arxiv.org/abs/2608.07082) |
| [Xtera_GSNR](./Xtera_GSNR/) | **Generalized SNR for Submarine Systems — Calculating performance in repeatered open line systems** | [Xtera Knowledge Centre](https://xtera.com/knowledge-centre/) |

각 하위 폴더의 `readme.md`에서 논문 정보, 재현 대상 그림, 모델 가정 및 실행 가능한 Colab 노트북을 확인할 수 있습니다.

## Horseshoe-and-Spur HCF 논문 설명

[Enhanced Scalability of Horseshoe-and-Spur Networks by Exploiting Hollow-Core Fiber](https://arxiv.org/abs/2608.07082)는 metro와 access 계층을 통합하는 filterless horseshoe-and-spur 광망에서 hollow-core fiber(HCF)와 DSCM 기반 coherent point-to-multipoint 송수신기를 함께 사용할 때의 확장성을 분석합니다. 두 hub 사이의 main horseshoe와 transit node에서 갈라지는 passive spur를 대상으로, 제한된 수의 광증폭기를 어디에 배치하고 coupler 비율과 spur 전력 예산을 어떻게 정해야 하는지를 최적화합니다.

기존 solid-core fiber(SCF)는 채널 출력을 높일 때 Kerr 비선형성이 먼저 제한하지만, HCF는 비선형성이 매우 작아 병목이 광증폭기의 총 출력으로 이동합니다. 논문에서는 이 차이로 HCF가 SCF보다 spur 전력 예산을 최대 약 20 dB 높일 수 있고, 5개 transit node 네트워크에서는 증폭기 수가 약 11–13개에 이르면 추가 설치 효과가 포화됨을 보입니다. 또한 기존 SCF horseshoe를 유지하면서 spur만 HCF로 교체하는 hybrid 구조가 full HCF와 SCF 사이의 성능을 제공하며, 줄어든 증폭기 비용으로 HCF의 추가 광섬유 비용을 상쇄할 가능성을 제시합니다.

`Horseshoe` 프로젝트는 논문 그래프의 좌표를 수동 입력하지 않습니다. Fig. 2와 Fig. 5는 공개 수식으로 직접 계산하고, Fig. 3과 Fig. 4는 공개된 수신감도·전력 불균형·비선형 임계값·증폭기 총출력·coupler 조건을 반영한 재현 가능한 MILP 대리모델로 계산합니다. 세부 가정과 한계는 [Horseshoe README](./Horseshoe/readme.md), 실행 코드는 [Colab 노트북](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/paper/Horseshoe/horseshoe_fig2_fig3_fig4_fig5_reproduction.ipynb)에서 확인할 수 있습니다.

## Xtera Generalized SNR 백서 설명

[Generalized SNR for Submarine Systems — Calculating performance in repeatered open line systems](https://xtera.com/knowledge-centre/)는 cable supplier와 coherent terminal supplier가 분리되는 repeatered submarine open line system에서 전송 성능을 terminal-independent GSNR로 표현하는 방법을 설명합니다. 선형 ASE 잡음의 \(SNR_O\)와 Kerr 비선형 잡음의 \(SNR_{NL}\)를 분리하고, 두 역수의 합으로 \(SNR_G\)를 계산합니다.

백서의 핵심은 기준 channel plan에서 주파수별 total noise factor \(F_{total}(f)\)와 nonlinear coefficient \(K_{NL}(f)\)를 한 번 구한 뒤, 이를 다른 채널 수·spacing·baud rate·출력 tilt 조건에 재사용하는 것입니다. 25 × 140 km Raman repeatered link에서 100 × 50 GHz, 34 GBaud 기준 plan과 133 × 37.5 GHz, 67 × 75 GHz, 50 GBaud의 세 변경 사례를 비교하여, 채널당 출력·ASE noise bandwidth·비선형 power spectral density 사이의 trade-off를 보여줍니다.

[Xtera_GSNR](./Xtera_GSNR/) 프로젝트는 Fig. 2–5의 좌표를 수동 입력하지 않습니다. 백서 식 (3)–(7)로 channel power, 누적 noise factor, \(SNR_O\), \(SNR_{NL}\), \(SNR_G\)를 계산하고, 공개되지 않은 Raman·fiber 세부값은 대표 SMF 파라미터와 closed-form incoherent GN surrogate로 대체합니다. 세부 식, 결과 및 한계는 [Xtera_GSNR README](./Xtera_GSNR/readme.md), 실행 코드는 [Colab 노트북](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/paper/Xtera_GSNR/generalized_snr_fig2_fig5_reproduction.ipynb)에서 확인할 수 있습니다.

