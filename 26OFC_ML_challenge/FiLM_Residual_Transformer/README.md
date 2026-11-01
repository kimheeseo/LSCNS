# Residual-FiLM Transformer | SHB-aware EDFA Gain Prediction

**목적:** OFC 2026의 95채널 EDFA Gain Profile을 물리 근사 기준선과 조건부 Attention의 결합으로 예측하는 실험 모델입니다.

## 핵심 기술
1. **Residual Learning:** `최종 Gain = Analytical Baseline(Pin, Gain, Tilt) + Transformer Residual`
2. **FiLM Conditioning:** 실제 제공되는 Aging 정보, Target Gain, Tilt로 각 Transformer 블록의 특징을 조건부 조절
3. **SHB-aware Curvature Loss:** 실측 곡률보다 과도한 국소 진동에만 추가 패널티 부여
4. **COSMOS Transfer 옵션:** COSMOS 사전학습 후 최신 Train 미세조정 및 Scratch 비교

> Analytical Baseline은 경험적 근사이며 EDFA rate equation의 직접 해가 아닙니다. 공식 평가의 `Category=aging`을 추론용 Aging 입력으로 사용하지 않습니다.

## 예비 성능 (2026-11-02)
**COSMOS 12,000개 부분집합 / 최대 18 epoch 사전학습**, 최신 Train 150 epoch, seed 2026, LR 0.002, curvature 0.02의 **사후 공개 TestGround** 결과입니다. *점수는 낮을수록 좋습니다.*

| 학습 방식 | 종합 | Aging | SHB | Unseen |
|---|---:|---:|---:|---:|
| COSMOS→Train | 1.572768 | 1.314756 | 1.447053 | 1.608132 |
| Train only | **1.540440** | 1.342096 | 1.539642 | 1.580112 |

**확인된 과제:** 내부 검증에서 SHB 극단 오차 발생. 기존 모델에 사용된 입력·출력 총파워 특징이 현재 모델에서 빠져 있어, 이를 복구하고 curvature ablation을 수행해야 합니다. **현 단계에서 정확도 개선을 주장할 수 없습니다.**

## 주요 파일
- [model.py](model.py) — Baseline, FiLM 블록, Residual 모델, SHB 곡률 손실
- [data.py](data.py) — 95채널 입력, Mask 및 Aging 컬럼 검증
- [train.py](train.py) — 학습·미세조정·평가·제출 CSV 생성
- [test_pipeline.py](test_pipeline.py) — 합성 데이터 기반 기능 테스트
- [Colab 노트북](OFC2026_Residual_FiLM_Transformer_Colab.ipynb) · [실험 기록](benchmarks/2026-11-02_cpu/README.md)

## 실행
```bash
pip install torch numpy pandas scikit-learn matplotlib
python test_pipeline.py
python train.py --data-root /path/to/ofc-ml-challenge-data-code \
  --output results/train_only --epochs 200 --lr 0.002
```

전이학습: `--pretrain-cosmos --cosmos-max-rows 4000 --scratch-control` 추가. 공식 [데이터 및 코드](https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code) 필요.

**상태:** 구조 구현 및 예비 학습 완료, 입력 특징 복구와 SHB 강건성 개선 필요.
