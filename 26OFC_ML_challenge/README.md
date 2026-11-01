# OFC 2026 ML Challenge | EDFA Gain Profile Prediction

**연구 목표:** 물리 모델, Spectral Attention, FiLM 및 COSMOS 전이학습을 비교하여 **95채널 EDFA Gain Profile** 예측의 일반화 성능을 개선합니다.

## 연구 모듈 안내

| 분류(권장 표시명) | 실제 GitHub 폴더 | 핵심 기능 |
|---|---|---|
| 01 · COSMOS Pretrain & Finetune | [COSMOS_Transfer_Learning](COSMOS_Transfer_Learning/) | PI-EDFA·Transformer COSMOS 사전학습 및 Scratch 대조 |
| 02 · Residual-FiLM SHB Transformer | [FiLM_Residual_Transformer](FiLM_Residual_Transformer/) | 기준선+FiLM 조건부 Attention+SHB 곡률 손실 |
| 03 · Spectral Attention Transformer | [PI_EDFA_Transformer](PI_EDFA_Transformer/) | 95채널 관계 학습과 Gain 잔차 예측 |
| 04 · Physics-Guided EDFA Benchmark | [results_2026-11-02](results_2026-11-02/) | PI-EDFA 공식 데이터 사후 검증 기록 |

> 표의 01~04는 **기능별 권장 이름**이며, 기존 Colab 및 상대경로 링크의 호환성을 위해 **실제 폴더 경로는 변경하지 않았습니다.**

## 확인된 성능 (2026-11-02)

### 동일 조건 공개 TestGround 사후 비교
*공식 공개 TestGround 21,056개 샘플; 낮을수록 좋습니다.*

| 모델 | Aging | SHB | Unseen | 종합 |
|---|---:|---:|---:|---:|
| PI-EDFA | **0.144064** | **0.147435** | 0.440072 | 0.382980 |
| Spectral Transformer | 0.163582 | 0.193067 | **0.364862** | **0.333409** |

- Spectral Transformer는 **Unseen·전체 점수 개선**, Aging·SHB는 악화
- 해당 점수는 공개 데이터 **사후 평가**로 비공개 대회 리더보드 점수가 아님

### 별도 예비 실험 (직접 순위 비교 금지)
- **COSMOS Transfer:** 4,000개 COSMOS 표본·6 epoch·동일 LR 실험에서 PI-EDFA 내부 검증 0.219143(Transfer) vs 0.231684(Scratch). Transformer는 0.302495 vs 0.274992
- **Residual-FiLM:** 12,000개 COSMOS 부분집합 예비 실험의 공개 TestGround 종합 1.572768(Transfer) / 1.540440(Scratch). SHB 이상치 및 누락된 전력 입력 특성 개선 필요

**주의:** 예비 실험의 데이터 범위·모델·분할·학습 조건이 다르므로 위 값을 하나의 순위로 취급하지 않습니다. TestGround 정답은 학습 및 모델 선택에 사용하지 않습니다.

## 실행 자료
- [PI-EDFA 원본 Colab](OFC_2026_Physics_Guided_EDFA_Colab.ipynb)
- [공식 측정 데이터 Benchmark Colab](OFC_2026_Official_Measured_Benchmark_Colab.ipynb)
- [공식 Subset Benchmark 스크립트](official_json_subset_benchmark.py)
- [공식 데이터 저장소](https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code)

모델별 설치·학습 명령, 입력 CSV, 검증 원칙은 해당 폴더의 `README.md`를 참조하세요.
