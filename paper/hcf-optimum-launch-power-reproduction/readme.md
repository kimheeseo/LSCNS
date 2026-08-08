# HCF Optimum Launch Power Reproduction

## Paper

**On the Optimum Energy-per-bit Launch Power in Coherent Hollow-core Fibre Transmission Systems**

- Authors: Ronit Sohanpal, Eric Sillekens, Mindaugas Jarmolovicius, Robert I. Killey, and Polina Bayvel
- Source: [arXiv:2606.17942](https://arxiv.org/abs/2606.17942)

이 디렉터리는 논문의 **Fig. 1, Fig. 2 및 Fig. 3(b)**를 Python/Google Colab으로 재현한 결과를 정리합니다. 그래프 좌표를 단순 입력하는 방식이 아니라, 논문에 제시된 광섬유 및 전송 시스템 파라미터를 기반으로 계산합니다.

## 논문 3줄 요약

1. 본 논문은 코히어런트 HCF 전송 시스템에서 최대 처리량 운용점과 최소 비트당 에너지 운용점이 서로 크게 달라질 수 있음을 분석합니다.
2. 파장별 HCF 손실·분산과 ASE 잡음, Kerr 비선형 간섭, 모드 간 간섭 및 트랜시버 잡음을 포함한 GN 기반 모델로 C-band와 초광대역 시스템의 처리량 및 에너지 효율을 평가합니다.
3. 1,000 km C-band HCF에서는 최대 처리량 대비 약 2.2%의 처리량 감소로 총 전력 소비를 41.5%, 비트당 에너지를 40.2% 절감할 수 있음을 제시합니다.

## Three-line Summary

1. This paper shows that the launch power for maximum throughput can differ substantially from that for minimum energy per bit in coherent HCF transmission systems.
2. A GN-based model incorporating wavelength-dependent HCF loss and dispersion, ASE noise, Kerr nonlinear interference, intermodal interference, and transceiver noise is used to evaluate C-band and ultra-wideband transmission.
3. For a 1,000 km C-band HCF link, operating with only a 2.2% throughput reduction can reduce total power consumption by 41.5% and energy per bit by 40.2% relative to maximum-throughput operation.

## Figure Reproduction Notebooks

| Figure | Notebook | 1줄 설명 |
|---|---|---|
| Fig. 1 | [dnanf_fig1_physics_model_colab.ipynb](./dnanf_fig1_physics_model_colab.ipynb) · [Open in Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/paper/hcf-optimum-launch-power-reproduction/dnanf_fig1_physics_model_colab.ipynb) | 해석적 DNANF 물리 모델을 이용하여 O/E/S/C/L-band 범위의 파장별 attenuation 및 dispersion profile을 계산합니다. |
| Fig. 2 | [dnanf_fig2_hcf_colab.ipynb](./dnanf_fig2_hcf_colab.ipynb) · [Open in Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/paper/hcf-optimum-launch-power-reproduction/dnanf_fig2_hcf_colab.ipynb) | ASE, closed-form GN 비선형 간섭, IMI 및 트랜시버 잡음을 계산하여 1×200 km C-band HCF의 launch power별 throughput을 재현합니다. |
| Fig. 3(b) | [dnanf_fig3b_hcf_colab.ipynb](./dnanf_fig3b_hcf_colab.ipynb) · [Open in Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/paper/hcf-optimum-launch-power-reproduction/dnanf_fig3b_hcf_colab.ipynb) | Fig. 2 처리량 모델에 증폭기 PCE와 채널당 24 W 트랜시버 전력을 결합하여 energy per bit와 throughput의 관계 및 최적 운용점을 계산합니다. |

## Notes

- 각 노트북은 Google Colab에서 독립적으로 실행할 수 있습니다.
- 논문 PDF의 벡터 곡선에서 추출한 비교 마커는 재현 정밀도 검증에만 사용되며, 물리 모델 계산값을 생성하거나 보정하는 입력으로 사용하지 않습니다.
- 원 논문의 full-vector FEM 데이터와 공개되지 않은 저출력 증폭기 PCE 피팅 계수는 재현 가능한 해석 모델 또는 명시적인 보정 계수로 대체했습니다.
