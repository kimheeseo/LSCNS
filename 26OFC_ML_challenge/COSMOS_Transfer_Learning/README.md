# COSMOS → Recent Train 전이학습 (PI-EDFA / Transformer)

기존 OFC 2026 모델 두 개를 **구조 변경 없이** 재사용하는 2단계 전이학습 실험입니다.

- 원본 PI-EDFA: [OFC_2026_Physics_Guided_EDFA_Colab.ipynb](../OFC_2026_Physics_Guided_EDFA_Colab.ipynb)
- 원본 Transformer: [transformer_edfa.py](../PI_EDFA_Transformer/transformer_edfa.py)
- 이 폴더의 [physics_edfa.py](physics_edfa.py)는 원본 노트북의 `%%writefile` 셀을 분리한 동일한 Python 모델 코드입니다.
- [train_transfer.py](train_transfer.py): 두 모델 공통 COSMOS 사전학습 → 최근 Train 미세조정 → 동일 검증셋 Scratch 대조 → 선택적 TestGround 평가
- [COSMOS_Transfer_Colab.ipynb](COSMOS_Transfer_Colab.ipynb): Colab에서 저장소·Git LFS 데이터 준비 후 두 모델 실행
- [COSMOS_Transfer_Error_Analysis_Report.ipynb](COSMOS_Transfer_Error_Analysis_Report.ipynb): **2026-11-02 실험 결과, 점수 보고 오류의 원인·수정, SHB 이상 샘플 ID, 입력 분포 이동, 학습률 통제 재실험 및 향후 검증계획을 기록한 연구 보고서**

## 데이터

[공식 OFC 2026 데이터 저장소](https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code)의 다음 파일을 사용합니다.

```text
Features/Train/COSMOS_features.csv       # 과거 대규모 입력 (Git LFS, ~110.5 MB)
Features/Train/COSMOS_labels.csv         # 과거 대규모 라벨 (Git LFS, ~56.1 MB)
Features/Train/train_features.csv        # 최근 Train 입력
Features/Train/train_labels.csv          # 최근 Train 라벨
Features/Test/test_features.csv          # 레이블 비공개 예측용
Features/TestGround/test_features.csv    # (선택) 대회 종료 후 사후 평가
Features/TestGround/test_labels.csv      # (선택) 대회 종료 후 사후 평가
```

Git LFS 포인터 파일이 아닌 실제 CSV를 받아야 합니다. `COSMOS_*.csv`를 Train과 단순 병합하는 것이 아니라 **별도 사전학습을 먼저 수행한 뒤 COSMOS의 가중치에서 미세조정**합니다. 모든 학습은 95채널 WSS 활성 마스크의 유효 채널만 손실 함수에 포함합니다.

## 학습 단계

1. COSMOS 데이터를 GroupShuffleSplit으로 분할하고, **COSMOS 학습 구간에만 fit한 scaler**로 모델의 입력을 표준화합니다.
2. COSMOS 학습으로 PI-EDFA / Transformer 각각의 best checkpoint를 저장합니다.
3. 최신 Train 데이터도 사전 결정된 동일한 group train/validation 분할을 사용합니다.
4. COSMOS 가중치와 **같은 scaler를 유지한 채** 작은 학습률(기본 1e-4)로 전 층 미세조정합니다.
5. 공정한 비교를 위해 최신 Train만으로 처음부터 학습하는 **scratch 모델**을 같은 분할로 따로 학습합니다.
6. 최근 Train validation에서 점수·MAE·카테고리별 점수를 비교합니다. `--evaluate-testground`는 모든 학습/모델 선택이 끝난 후 사후 보고용으로만 사용합니다.

**중요:** validation 점수는 조기 종료·best checkpoint 선택에 사용되므로 독립 TestGround 점수와 다릅니다. TestGround 정답을 학습, 스케일러 계산, 에폭 선택, 하이퍼파라미터 튜닝에 사용하면 안 됩니다.

## Colab 실행

[Colab 노트북](COSMOS_Transfer_Colab.ipynb)을 열고 위에서부터 실행하면 됩니다. T4 GPU 사용을 권장하며, 실제 COSMOS CSV 다운로드와 가중치 학습은 Colab 세션에서 이뤄집니다. GitHub에는 훈련 완료된 모델 가중치를 포함하지 않았습니다.

로컬 예제:

```bash
pip install torch numpy pandas scikit-learn matplotlib
python train_transfer.py \
  --data-root /path/to/ofc-ml-challenge-data-code \
  --output ./cosmos_results \
  --models pi-edfa transformer \
  --pretrain-epochs 60 --finetune-epochs 180
```

- 전체 COSMOS 학습: `--cosmos-max-rows 0` (기본값)
- 빠른 파일/파이프라인 확인: `--cosmos-max-rows 3000 --pretrain-epochs 2 --finetune-epochs 2` (**성능 평가용이 아님**)
- 성능 비교용 scratch 생략: `--skip-scratch`
- 대회 종료 후 공개 TestGround를 **사후에만** 확인: `--evaluate-testground`
- CPU 강제: `--cpu`

## 결과

```text
cosmos_results/
  comparison.json
  comparison.csv
  pi-edfa/
    cosmos_pretrain.pt
    recent_finetune.pt
    recent_scratch.pt
    cosmos_pretrain_history.csv
    recent_finetune_history.csv
    recent_scratch_history.csv
    results.json
    submission.csv
  transformer/
    (동일한 파일 구조)
```

`results.json`에 pretrain source validation, 최근 Train transfer validation, scratch validation 결과를 별도 저장합니다. `--evaluate-testground` 실행 시 `*_testground.json`이 추가됩니다. 점수는 작을수록 좋습니다.

## 이전 실험과의 차이

2026-11-02 원본 실험에서는 **COSMOS 사전학습 없이** PI-EDFA 0.382980, Transformer 0.333409의 공개 TestGround 전체 점수를 얻었습니다(대회 종료 후 사후평가). **이 숫자들은 새 전이학습의 결과가 아닙니다.** 이번 코드는 아직 실제 COSMOS 전체 데이터로 학습 실행·성능 측정되지 않았으며, Colab 실행 결과가 생성되어야 개선 여부를 결론낼 수 있습니다.

데이터와 모델 구조가 서로 호환되더라도 COSMOS→최근 Train 간 장비·시간·목표 gain/tilt 분포가 다르면 음의 전이가 발생할 수 있으므로 scratch 대조를 반드시 확인하세요.

## 2026-11-02 실제 COSMOS 예비 실험 (정식 전체 데이터 실험 아님)

공식 Git LFS의 COSMOS 4,000개 학습 표본, 6 epoch 사전학습 및 최신 Train 397/93 분할을 사용한 GitHub Actions 결과입니다. 검증 점수는 **최고 checkpoint 기준**이며 낮을수록 좋습니다.

| 모델 | Transfer + LR=0.002 | Scratch + LR=0.002 | 결과 |
|---|---:|---:|---|
| PI-EDFA | **0.219143** | 0.231684 | Transfer 5.4% 개선 |
| Transformer | 0.302495 | **0.274992** | Transfer 10.0% 악화 |

기존 LR=0.0001 Transfer에서는 SHB 극단 오차 약 29 dB로 점수가 3.6~3.7까지 악화되었습니다. 동일 LR=0.002 대조실험에서는 해당 극단오차가 해소되었습니다. **저학습률로 인한 미세조정 적응 부족**이 강하게 의심되지만, 공식 TestGround 재평가가 아니며 전체 COSMOS 사전학습 결론은 아닙니다.

오류 감사 기록: [연구 보고서 노트북](COSMOS_Transfer_Error_Analysis_Report.ipynb) · [최초 pilot](benchmarks/pilot/README.md) · [LR 통제 진단](benchmarks/lr_control/README.md)
