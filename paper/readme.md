# Paper Reproduction

이 폴더는 논문에 제시된 수치, 알고리즘 및 그래프를 Python/Google Colab으로 재검증한 프로젝트를 모아두기 위한 공간입니다.

## Reproduction Projects

| Folder | Based on | Source |
|---|---|---|
| [hcf-optimum-launch-power-reproduction](./hcf-optimum-launch-power-reproduction/) | **On the Optimum Energy-per-bit Launch Power in Coherent Hollow-core Fibre Transmission Systems** | [arXiv:2606.17942](https://arxiv.org/abs/2606.17942) |
| [opticalfiber](./opticalfiber/) | **High Density Optical Cable with Ultra-Low-Loss, Large-Effective-Area ITU-T G.654.E Optical Fiber** | [IWCS Webinar 85](https://iwcs.org/webinar/85/) |
| [EDFA-Raman-320-km](./EDFA-Raman-320-km/) | **320㎞ Optical Transmission using EDFA and Raman amplifier for 10Gbit/s 128 Channel DWDM Signals** | [KICS Digital Library](https://journal.kics.or.kr/digital-library/33081) |
| [Horseshoe](./Horseshoe/) | **Enhanced Scalability of Horseshoe-and-Spur Networks by Exploiting Hollow-Core Fiber** | [arXiv:2608.07082](https://arxiv.org/abs/2608.07082) |

각 하위 폴더의 `readme.md`에서 논문 정보, 재현 대상 그림, 모델 가정 및 실행 가능한 Colab 노트북을 확인할 수 있습니다.

## Horseshoe-and-Spur HCF 논문 설명

[Enhanced Scalability of Horseshoe-and-Spur Networks by Exploiting Hollow-Core Fiber](https://arxiv.org/abs/2608.07082)는 metro와 access 계층을 통합하는 filterless horseshoe-and-spur 광망에서 hollow-core fiber(HCF)와 DSCM 기반 coherent point-to-multipoint 송수신기를 함께 사용할 때의 확장성을 분석합니다. 두 hub 사이의 main horseshoe와 transit node에서 갈라지는 passive spur를 대상으로, 제한된 수의 광증폭기를 어디에 배치하고 coupler 비율과 spur 전력 예산을 어떻게 정해야 하는지를 최적화합니다.

기존 solid-core fiber(SCF)는 채널 출력을 높일 때 Kerr 비선형성이 먼저 제한하지만, HCF는 비선형성이 매우 작아 병목이 광증폭기의 총 출력으로 이동합니다. 논문에서는 이 차이로 HCF가 SCF보다 spur 전력 예산을 최대 약 20 dB 높일 수 있고, 5개 transit node 네트워크에서는 증폭기 수가 약 11–13개에 이르면 추가 설치 효과가 포화됨을 보입니다. 또한 기존 SCF horseshoe를 유지하면서 spur만 HCF로 교체하는 hybrid 구조가 full HCF와 SCF 사이의 성능을 제공하며, 줄어든 증폭기 비용으로 HCF의 추가 광섬유 비용을 상쇄할 가능성을 제시합니다.

`Horseshoe` 프로젝트는 논문 그래프의 좌표를 수동 입력하지 않습니다. Fig. 2와 Fig. 5는 공개 수식으로 직접 계산하고, Fig. 3과 Fig. 4는 공개된 수신감도·전력 불균형·비선형 임계값·증폭기 총출력·coupler 조건을 반영한 재현 가능한 MILP 대리모델로 계산합니다. 세부 가정과 한계는 [Horseshoe README](./Horseshoe/readme.md), 실행 코드는 [Colab 노트북](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/paper/Horseshoe/horseshoe_fig2_fig3_fig4_fig5_reproduction.ipynb)에서 확인할 수 있습니다.
