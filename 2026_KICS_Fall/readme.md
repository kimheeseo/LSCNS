# 2026 KICS Fall — WDM GN/EGN 분석

2026년 KICS 추계학회 연구를 위한 **장거리 WDM 광전송 비선형 성능 예측·검증 코드**입니다. 광섬유·WDM·증폭기 조건으로 NLI, ASE, GSNR 및 채널당 전송 성능을 계산합니다.

## 구성

| 위치 | 용도 |
|---|---|
| `gn_integral_general.py` | Sobol QMC 기반 GN 적분으로 NLI PSD와 CUT 수신대역 NLI 전력 계산 |
| `gn_integral_general_modulation.py` | ASE, GSNR, 근사 BER, 전송률·Shannon-gap 용량 계산 |
| `run_gn.py` | GN 계산·입사전력 sweep 사용 예제 |
| [conditions/](./conditions/) | KICS 재현용 기준 입력 조건 |
| [GNmodel/](./GNmodel/) | GN 수식·수치수렴·문헌 결과 검증 자료 |
| [EGN_model/](./EGN_model/) | Full-EGN 계산기 및 GN/EGN·문헌 검증 자료 |

## KICS 재현 기준

공식 조건은 [conditions/README.md](./conditions/README.md)를 따릅니다. 요약하면 50채널, 95 GBd/95 GHz, 0 dBm/ch, 80 km × 30 spans(2,400 km), 1550 nm 조건에서 계산합니다.

NLI는 Sobol seed 1·2·3을 각각 계산한 뒤 **선형 전력 단위(W)로 평균**합니다. CUT는 0-based index 24이며, GSNR 계산에는 TRX SNR 18 dB를 포함합니다.

> `run_gn.py`의 현재 값은 사용법을 보이기 위한 예제입니다. KICS 결과를 재현할 때는 채널·광섬유·누적 방식·Sobol 설정을 `conditions/`의 기준값으로 맞춰야 합니다.

## 실행

```bash
pip install numpy scipy
python run_gn.py
```

주요 함수는 `integrate_nli_over_channel()`, `evaluate_performance()`, `launch_power_vs_gsnr()`입니다.

## 해석 범위

- `nli_model="gn"`: 범용 GN 모델
- `nli_model="egn_sci"`: SCI만 보정한 근사 모델이며 Full-EGN은 아님
- Full-EGN(SCI+XCI+MCI)은 [EGN_model/EGN_adaptive.py](./EGN_model/EGN_adaptive.py)를 사용

ASE는 lumped-EDFA 조건을 가정합니다. BER·용량은 시스템 수준의 근사값이며, PMD/PDL, 레이저 위상잡음, 상세 DSP 및 ROADM 필터 누적 효과는 포함하지 않습니다.
