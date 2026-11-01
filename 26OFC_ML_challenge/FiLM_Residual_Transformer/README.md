# OFC 2026 Residual-FiLM Transformer (95-channel EDFA gain)

기존 `26OFC_ML_challenge/PI_EDFA_Transformer`를 유지한 채 만든 **독립 실행형 PyTorch 구현**입니다. OFC 2026 공개 CSV 기준으로 학습, 검증, 제출 형식 저장, 선택적 COSMOS 사전학습, 학습률 통제 Scratch 대조실험이 가능합니다.

## 3개 구조적 개선

1. **Residual Learning**: `Final Gain = AnalyticalGainBaseline(Pin, Gain, Tilt) + learned residual`. 기준선은 Gain/Tilt + 입력 스펙트럼 중심화 편차의 경험적 보정이며, 에르븀 rate equation의 해가 **아닙니다**.
2. **FiLM in each Transformer block**: Age(관측되는 경우), Target Gain, Target Tilt를 각 블록의 Attention·FFN 경로에 `FiLM(h,c) = (1+gamma(c))h + beta(c)` 형태로 적용합니다. Token에 스칼라 값을 직접 concat하지 않습니다. 각 FiLM은 identity로 초기화합니다.
3. **SHB-aware Curvature Loss**: `MSE_active + lambda * mean( relu(|D² y_pred| - |D² y_true| - tolerance)^2 )`이며, 곡률 항은 세 채널이 연속 활성일 때만 적용됩니다. 실제 측정된 SHB 커브의 곡률까지 일률적으로 0에 수렴시키지 않습니다.

**주의:** 일반적 smoothness prior는 실제 국소 SHB 홀을 뭉개는 위험이 있으므로 이 설계는 **예측이 실측 곡률보다 과도한 정도**를 벌점으로 줍니다. 이 방식만으로 모든 고주파 노이즈가 제거된다는 보장은 없으며 `--curvature 0,0.005,0.02,0.05` ablation을 권장합니다.

## 파일

| 파일 | 역할 |
|---|---|
| `model.py` | Baseline, FiLM, FiLM Transformer blocks, ResidualFiLMEDFA, SHB-aware loss, single train step |
| `data.py` | 공식 데이터의 95채널 컬럼 읽기, mask/label ID 정합성, Group split, age 제어변수 검증 |
| `train.py` | 2단계 COSMOS→Train 학습, Train-only 대조, best checkpoint, OFC score, TestGround 사후 감사 및 `submission.csv` |
| `test_pipeline.py` | 합성 데이터로 forward/backward, NaN mask, checkpoint 복원, end-to-end CSV 실행 테스트 |
| `OFC2026_Residual_FiLM_Transformer_Colab.ipynb` | GPU Colab에서 원클릭 데이터 준비·학습·평가·그래프·ZIP 저장 |

## 실행

Python 3.11 권장. Colab T4 GPU에서 실행하려면 [노트북](OFC2026_Residual_FiLM_Transformer_Colab.ipynb)을 열어 셀 순서대로 실행합니다.

```bash
pip install torch numpy pandas scikit-learn matplotlib
python test_pipeline.py

# 공식 Train만 사용하여 학습
python train.py --data-root /path/to/ofc-ml-challenge-data-code \
  --output results/train_only --epochs 200 --lr 0.002

# COSMOS 사전학습 → Train 미세조정 + 같은 Split Scratch 대조
python train.py --data-root /path/to/ofc-ml-challenge-data-code \
  --output results/cosmos_transfer --pretrain-cosmos --cosmos-max-rows 4000 \
  --pretrain-epochs 40 --epochs 200 --lr 0.002 --scratch-control

# 전체 COSMOS 사용: --cosmos-max-rows 0

# 모델 확정 후 공개 TestGround 사후 검증 (학습하지 않음)
python train.py --data-root /path/to/ofc-ml-challenge-data-code \
  --output results/final_audit --audit-checkpoint results/cosmos_transfer/recent_best.pt
```

공식 데이터: https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code

필요한 Git LFS 파일:

```text
Features/Train/train_features.csv
Features/Train/train_labels.csv
Features/Train/COSMOS_features.csv        # --pretrain-cosmos 사용할 때
Features/Train/COSMOS_labels.csv          # --pretrain-cosmos 사용할 때
Features/Test/test_features.csv           # 제출 CSV 만들 때
Features/TestGround/test_features.csv     # frozen audit에만
Features/TestGround/test_labels.csv       # frozen audit에만
```

## Age / Aging 정보에 관한 데이터 누출 방지

공식 `Category=aging`은 평가 데이터의 분류 라벨입니다. 모델의 추론 시점에 실제 소자 노후도가 제공된다는 보장이 없으므로 **Category를 FiLM의 Age 값으로 사용하지 않습니다.** `aging_hours`, `aging_days`, `age_hours`, `age_days`, `EDFA_age_hours`, `device_age_hours`, `DUT_age`, `age` 등 **입력 features CSV에 실제 측정 수치가 있으면** 사용하고, 없으면 Age를 0으로 고정합니다. 외부에서 정당하게 사용할 수 있는 노후도 컬럼이 있다면 `--age-column <name>`으로 명시합니다. `Category`, `ID`, 예측 라벨은 거부합니다. Age가 상수라면 Age FiLM은 작동하더라도 노후도 효과를 식별 학습할 수 없습니다.

조건 변수의 평균·표준편차는 반드시 학습 구간에서만 fit하며, COSMOS→Train 미세조정에서는 사전학습 때 사용한 조건 정규화를 **변경하지 않습니다**. 최신 Train 검증과 공개 TestGround 정답은 학습에 사용하지 않습니다.

## 결과와 해석

- `recent_best.pt`, `recent_history.csv`, `recent_metrics.json`, `recent_per_sample.csv`
- COSMOS 실행 시 `cosmos_pretrain_best.pt`와 관련 학습 기록
- `--scratch-control` 사용 시 `scratch_best.pt`, `scratch_metrics.json`
- `comparison.json`: 모델 선택 시점의 **best checkpoint** 점수만 저장
- `submission.csv`: `ID` + 95 Gain 채널 컬럼. WSS OFF 채널은 -18 dB
- `frozen_checkpoint_testground_audit.json`: 별도 감사 실행 시 생성

이 프로그램은 3개 구조적 개선을 구현한 연구용 코드이며, **기존 Transformer보다 점수가 개선된다는 보장은 없습니다.** 특히 공개 TestGround 점수와 내부 Train 검증 점수를 직접 비교할 수 없습니다. 동일 분할, seed, epoch, learning rate, label mask 조건에서 `curvature=0` 대비 Ablation Test가 필요합니다.

## 개발 검증

`python test_pipeline.py`의 숫자는 임의 생성한 합성 데이터의 **기능 테스트용** 지표일 뿐 OFC 대회 점수가 아닙니다. 공개 데이터에서 충분히 학습한 모델의 실제 Accuracy는 학습 수행 후 확인해야 합니다.