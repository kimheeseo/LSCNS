# PI-EDFA Transformer — COSMOS 사전학습 및 대회 데이터 미세조정

기존 `../PI_EDFA_Transformer/transformer_edfa.py`의 PhysicalSpectralNet·물리 기반 입력·점수 대리손실을 재사용하는 실험 코드입니다. 우승자 모델·가중치를 복사하지 않습니다. 기존 모델에 COSMOS 사전학습 → 대회 데이터 미세조정 단계를 추가했습니다.

## 실행 조건

Python 3, torch, numpy, pandas, scikit-learn, matplotlib 및 기존 PI_EDFA_Transformer 폴더가 필요합니다. COSMOS와 대회 데이터는 모두 기존 모델과 같은 95채널 CSV 스키마여야 하며, COSMOS 원시 JSON을 직접 변환하지 않습니다. 대회 규칙과 데이터 사용 조건을 확인하세요. Colab T4 또는 CPU에서 실행 가능합니다.

```bash
pip install torch numpy pandas scikit-learn matplotlib
python 26OFC_ML_challenge/PI_EDFA_Transformer_TwoStage/train_two_stage.py \
 --cosmos-features /path/COSMOS_features.csv \
 --cosmos-labels /path/COSMOS_labels.csv \
 --train-features /path/train_features.csv \
 --train-labels /path/train_labels.csv \
 --test-features /path/test_features.csv \
 --output /path/results
```

Drive를 마운트한 Colab은 각 경로에 `/content/drive/MyDrive/...`를 지정하세요. --cpu 옵션으로 CPU 강제, --pretrain-epochs 및 --finetune-epochs로 단계별 학습량 조정이 가능합니다. 그룹이 충분하면 기존 모델의 get_groups 기준 분할, 그렇지 않으면 랜덤 분할을 하고 metrics.json에 기록합니다.

## 산출물

- submission.csv: 테스트 입력 기반 95채널 예측. 실제 대회 sample_submission과 컬럼·ID·비활성 값 일치 여부를 제출 전에 확인하세요.
- model.pt: 가중치 및 표준화 평균·스케일(optimizer/재개 상태는 미포함).
- metrics.json: 내부 검증점수·MAE·모델 크기·추론 시간·분할 방식.
- history.csv, split_indices.npz: 학습 기록과 분할 인덱스. COSMOS 인덱스는 중복 제거 후 데이터 기준입니다.

전처리기는 COSMOS 훈련과 대회 훈련 데이터에만 맞추고 검증과 테스트는 제외합니다. 공통 입력 열의 정확히 동일한 행은 COSMOS에서 제거하나, 측정 시점·장비 연관 및 유사 샘플 누출까지 배제하지 못합니다. 개별 소스 파일의 정렬·기기 식별·데이터 허용 조건을 확인해야 합니다. 비활성 채널은 손실에서 제외합니다.

## 상태와 한계

**문법 검사는 완료했으나 실제 데이터 학습·추론·대회 제출 검증은 아직 수행하지 않았습니다.** 기존 score 함수가 대회 공식 평가와 일치하는지 및 test_features와 sample_submission 규격, 출력파워 특성의 실제 추론 시 가용성을 확인해야 합니다. Aging/SHB/Unseen 별도 평가, 체크포인트 학습 재개, Colab 노트북 및 다중 seed 성능 검증은 아직 구현하지 않았습니다. 우승 수준의 정확도를 보장하지 않습니다. 보고된 다른 데이터셋 점수와 직접 비교하지 마세요.
