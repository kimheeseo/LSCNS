# COSMOS Pretraining & Fine-tuning | PI-EDFA + Transformer

**목적:** OFC 2026 EDFA 95채널 Gain Profile 예측에서 COSMOS 사전학습이 최신 Train 일반화 성능을 개선하는지 검증합니다.

## 핵심 구조
- **2단계 학습:** COSMOS 사전학습 → 최신 Train 미세조정
- **비교 모델:** 물리 기반 PI-EDFA와 Spectral Transformer 각각에 동일 학습 절차 적용
- **대조 실험:** 최신 Train만 학습한 Scratch 모델과 동일 분할·학습률 조건에서 비교
- **평가 원칙:** WSS 활성 채널만 점수 계산. TestGround는 모델 확정 후 사후 평가에만 사용

## 예비 결과 (2026-11-02)
COSMOS **4,000개 표본 / 6 epoch**, Transfer와 Scratch **LR=0.002** 통제 실험의 **최적 체크포인트 기준 Train 내부 검증 점수**입니다. *낮을수록 좋습니다.*

| 모델 | COSMOS→Train | Train only | 해석 |
|---|---:|---:|---|
| PI-EDFA | **0.219143** | 0.231684 | 전이학습 개선 |
| Transformer | 0.302495 | **0.274992** | 전이학습 악화 |

> 전체 COSMOS 학습 결과 또는 공개 TestGround 점수가 아닙니다. 사전학습 효과는 모델마다 다르며, 학습률에 민감합니다.

## 주요 파일
- [train_transfer.py](train_transfer.py) — 사전학습, 미세조정, Scratch 비교 및 평가
- [physics_edfa.py](physics_edfa.py) — 재사용하는 PI-EDFA 모델 구현
- [COSMOS_Transfer_Colab.ipynb](COSMOS_Transfer_Colab.ipynb) — Google Colab 실행
- [COSMOS_Transfer_Error_Analysis_Report.ipynb](COSMOS_Transfer_Error_Analysis_Report.ipynb) — 오류 분석 및 재실험 보고
- [LR 통제 실험](benchmarks/lr_control/README.md) · [최초 Pilot 기록](benchmarks/pilot/README.md)

## 실행
공식 [데이터 저장소](https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code)의 **Git LFS 실제 CSV**가 필요합니다.

```bash
pip install torch numpy pandas scikit-learn matplotlib
python train_transfer.py --data-root /path/to/ofc-ml-challenge-data-code \
  --output ./cosmos_results --models pi-edfa transformer \
  --pretrain-epochs 60 --finetune-epochs 180
```

- 빠른 기능 점검: `--cosmos-max-rows 3000 --pretrain-epochs 2 --finetune-epochs 2`
- 동일 LR 검증 시 `--finetune-lr 0.002 --scratch-lr 0.002` 명시
- 공개 TestGround 사후 확인: `--evaluate-testground` (훈련 및 모델 선택에 사용 금지)

**상태:** 소규모 사전학습 실험 확인, 전체 COSMOS 재현·독립 성능 검증 추가 필요.
