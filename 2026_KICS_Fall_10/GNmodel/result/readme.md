# GN 모델 검증 결과 파일

| 파일 | 설명 |
|---|---|
| [Carena_Fig5a_PSCF.png](./Carena_Fig5a_PSCF.png) · [PDF](./Carena_Fig5a_PSCF.pdf) | Carena 그림 5(a), PSCF의 저자 GN 실선(Paper)과 TXT 계산값(Code) 비교. |
| [Carena_Fig5b_SMF.png](./Carena_Fig5b_SMF.png) · [PDF](./Carena_Fig5b_SMF.pdf) | 같은 방식의 그림 5(b), SMF 비교. |
| [Carena_Fig5c_NZDSF.png](./Carena_Fig5c_NZDSF.png) · [PDF](./Carena_Fig5c_NZDSF.pdf) | 같은 방식의 그림 5(c), NZDSF 비교. |
| [PSCF_solid_vs_TXT.csv](./PSCF_solid_vs_TXT.csv) · [SMF](./SMF_solid_vs_TXT.csv) · [NZDSF](./NZDSF_solid_vs_TXT.csv) | 변조·채널 간격별 Paper/Code 거리, 최적 입사전력, NLI 계수 및 상대오차. |
| [Poggiolini_assessment.md](./Poggiolini_assessment.md) | GN 수식 검증 7개와 Sobol 적분 수렴 평가 보고서. |

Carena 비교는 시뮬레이션 마커가 아닌 **저자 GN 실선의 이미지 판독값**을 사용합니다. 각 21개 점의 MAPE는 PSCF 약 **5.4%**, SMF 약 **5.2%**, NZDSF 약 **9.8%**입니다. NZDSF의 100 km 계산점은 그래프 범위 아래에 있으며 CSV에는 포함합니다.

계산은 9채널·32 GBd·100 km 스팬·NF 5 dB·비코히어런트 GN, Sobol 2¹⁸·시드 2·수신 적분점 11개, TRX 잡음 제외 조건입니다. 그림 3의 요구 OSNR과 근사 송신 PSD, 코드의 직사각형 수신 대역을 사용합니다. 실선 판독·보간 및 수신기 근사가 있으므로 완전한 논문 시스템 재현이나 절대 정확도 보증은 아닙니다.

Poggiolini 평가의 **PASS**는 선정 수식의 구현 일관성 검증으로, 위 거리 MAPE와 다른 지표입니다.
