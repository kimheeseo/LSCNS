# EDFA/Raman 320 km Transmission Reproduction

## Paper

**320㎞ Optical Transmission using EDFA and Raman amplifier for 10Gbit/s 128 Channel DWDM Signals**

- Author: Bo-Hun Choi
- Journal: The Journal of Korean Institute of Communications and Information Sciences, Vol. 34, No. 6, 2009
- Source: [KICS Digital Library](https://journal.kics.or.kr/digital-library/33081)
- Original Korean title: **10 Gbit/s 128 채널 고밀도 파장다중화 신호를 위해 EDFA와 라만 증폭기를 이용한 320km 광전송 실험**

이 디렉터리는 논문의 **Fig. 3, Fig. 7, Fig. 9 및 Fig. 10**을 Python/Google Colab으로 재현하고 논문 결과와 비교한 내용을 정리합니다. 논문 그래프 좌표를 Python 모델의 입력값으로 직접 사용하지 않고, 논문에 공개된 식과 전송 조건 및 명시적인 공학적 가정을 이용해 결과를 계산합니다.

## 논문 3줄 요약

1. 본 논문은 C-band와 L-band에 각각 64채널을 배치한 총 128채널, 채널당 10 Gbit/s DWDM 신호의 320 km 전송을 위해 EDFA와 분산형 라만 증폭기를 함께 사용하는 링크를 설계하고 실험합니다.
2. 80 km NZ-DSF 구간 손실 22 dB, OBA·OLA·OPA 이득 20 dB 및 측정된 NZ-DSF 잡음지수 관계를 이용하여 증폭기 수와 라만 이득에 따른 OSNR을 계산하고, C/L-band 2단 EDFA를 제작합니다.
3. 제작된 링크의 320 km 전송 후 C-band와 L-band에서 평균 약 25 dB의 OSNR을 측정하여, 10 Gbit/s급 128채널 DWDM 전송 가능성을 확인합니다.

## Three-line Summary

1. This paper designs and experimentally demonstrates a 320 km link for 128-channel 10 Gbit/s DWDM transmission using dual-band EDFAs together with distributed Raman amplification.
2. The link OSNR is calculated from the measured NZ-DSF noise-figure relation, a 22 dB loss per 80 km span, and 20 dB OBA/OLA/OPA gains, after which two-stage C- and L-band EDFAs are implemented.
3. The experiment obtains an average OSNR of approximately 25 dB in both bands after 320 km, supporting the feasibility of the proposed 128-channel transmission system.

## 주요 논문 조건

| 항목 | 논문 조건 |
|---|---|
| 총 채널 수 | 128채널: C-band 64채널 + L-band 64채널 |
| 채널 전송속도 | 10 Gbit/s |
| 채널 간격 | 50 GHz, 약 0.4 nm |
| 목표 전송거리 | 320 km |
| 전송 광섬유 | NZ-DSF |
| 기본 구간 | 80 km, 손실 22 dB |
| 증폭기 이득 | OBA·OLA·OPA 각각 20 dB |
| 채널당 입력 | OBA −20 dBm, OLA·OPA −15 dBm |
| Fig. 3 EDFA 잡음지수 | 7 dB |
| Fig. 3 라만 이득 범위 | 0, 2, 3, 4, 5, 6 dB |
| L-band EDFA의 EDF 길이 | 1단 50 m, 2단 80 m |
| 측정 거리 | 0, 80, 140, 200, 260, 320 km |

## 그림별 의미

| Figure | 논문에서의 성격 | 주요 내용 |
|---|---|---|
| Fig. 3 | 컴퓨터 시뮬레이션 | 라만 이득과 증폭기 구간 수 변화에 따른 OSNR 계산 |
| Fig. 7 | 실험 측정 | 제작된 L-band 2단 EDFA의 파장별 이득과 잡음지수 |
| Fig. 9 | 실험 측정 | 0–320 km 거리별 C-band 채널 OSNR |
| Fig. 10 | 실험 측정 | 0–320 km 거리별 L-band 채널 OSNR |

## Figure Reproduction Notebooks

| Figure | Notebook | 1줄 설명 |
|---|---|---|
| Fig. 3 | [edfa_raman_fig3_reproduction_colab.ipynb](./edfa_raman_fig3_reproduction_colab.ipynb) · [Open in Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/paper/EDFA-Raman-320-km/edfa_raman_fig3_reproduction_colab.ipynb) | 논문의 식 (1) Friis 잡음지수와 식 (2) OSNR 관계 및 Fig. 2의 NZ-DSF 회귀식을 이용해 라만 이득별 OSNR을 계산합니다. |
| Figs. 3, 7, 9, 10 | [edfa_raman_fig3_fig7_fig9_fig10_reproduction_colab.ipynb](./edfa_raman_fig3_fig7_fig9_fig10_reproduction_colab.ipynb) · [Open in Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/paper/EDFA-Raman-320-km/edfa_raman_fig3_fig7_fig9_fig10_reproduction_colab.ipynb) | Fig. 3 링크 계산에 더해 2준위·McCumber EDFA 모델과 ASE 잡음 누적 모델로 Figs. 7, 9, 10을 계산하고 논문 측정값과 비교합니다. |

## 재현 알고리즘

- **Fig. 3:** 논문의 Friis cascade noise-factor 식과 OSNR 식, Fig. 2의 NZ-DSF 회귀식 `NF(dB) = -0.2025 × G(dB) + 17.08`을 이용합니다.
- **Fig. 7:** 분석적 erbium 흡수 단면적, McCumber 관계식, 2준위 균일 반전 모델 및 파장 의존 수동 손실을 이용해 L-band EDFA의 이득과 잡음지수를 계산합니다.
- **Figs. 9·10:** OBA, 등가 전송 구간, OLA 및 OPA에서 발생하는 ASE 잡음을 reciprocal-linear 방식으로 누적하여 거리별 OSNR 스펙트럼을 계산합니다.
- **논문값 비교:** 논문 PDF의 색상 곡선을 자동 추출한 값은 Python 모델 계산이 끝난 후 빈 원형 마커와 오차 계산에만 사용합니다.

## 현재 모델의 비교 오차

| 비교 항목 | 평균 절대오차(MAE) |
|---|---:|
| Fig. 7 gain | 약 1.08 dB |
| Fig. 7 noise figure | 약 1.00 dB |
| Fig. 9 C-band OSNR | 약 1.58 dB |
| Fig. 10 L-band OSNR | 약 1.08 dB |

## Notes

- 각 노트북은 Google Colab에서 독립적으로 실행할 수 있습니다.
- **Python model**은 알고리즘 계산 결과를 실선으로, **Paper measurement**는 논문 PDF에서 자동 추출한 비교값을 빈 원형 마커로 표시합니다.
- Fig. 3은 논문의 실제 컴퓨터 시뮬레이션 결과이며, Figs. 7, 9, 10은 실험 측정 결과입니다.
- 측정 그래프의 값을 모델에 대입하거나 파장별로 역보정하지 않습니다. 논문에 명시된 19.7 dB 평균 이득이나 25 dB 평균 OSNR 결과도 모델의 목표값으로 역산하지 않습니다.
- 논문에는 실제 라만 펌프 출력, 실험 링크의 on/off Raman gain, EDF 흡수·방출 단면적 원본, Er 이온 농도, AOTF 전달함수, 부품별 삽입손실 및 OSA 원시 데이터가 공개되어 있지 않습니다.
- 따라서 Figs. 7, 9, 10은 공개된 정보와 명시적인 공학적 가정에 기반한 독립 재현이며, 위 오차는 누락된 실험 파라미터에 따른 모델 불확실성을 포함합니다.
