# 첨부 학습 로그 진단 및 재개 수정

첨부 로그는 `resnet18/original`의 조기 종료 상태를 불러온 뒤 `resnet18/augmented`를 완료된 6 Epoch 다음인 7 Epoch부터 재개하여, 10 Epoch의 60/299 배치까지 진행한 기록이다. Traceback, OOM, NaN 기록은 없다. 이후 중단 여부와 원인은 이 기록만으로 확정할 수 없다.

## 진행하던 작업

광섬유 SEM PNG를 bending / cleaving / fracture 세 클래스로 분류하는 사전학습 CNN 비교이다. HCF는 구조 태그이지 별도 정답 클래스가 아니다.

1. Google Drive의 저장 이미지 4,876개를 SHA-256·PNG 디코딩 검증 후 재로딩했다. 10분 19초는 이 단계의 시간이다. 4,876개는 증강 파일을 포함하며 독립 원본 수가 아니다.
2. ResNet18 원본 전용 모델: 저장된 Early Stopping 상태이므로 추가 학습 없이 최적 모델을 평가했다.
3. ResNet18 증강 모델: 7~10 Epoch를 CPU, batch 16, 입력 384px로 학습 중이었다.
4. 이후 EfficientNet-B0 원본 전용 → 증강 모델 학습·평가, 네 모델의 보고서 집계 및 선택 모델 저장이 이어진다. 기존 체크포인트가 있으면 각 상태에 맞춰 재개/평가한다.
5. 독립 시료 그룹이 부족한 조건에서는 별도 그룹 교차검증도 실행된다. Validation은 모델 선택에도 사용되므로 독립 Test 점수가 아니다.

기존 배치 로그의 퍼센트는 **현재 모델의 최대 Epoch 대비** 값이다. 10/10 Epoch도 검증·체크포인트·최종 평가·특징 추출·보고서 작성이 남는다.

## 확인된 문제와 해석

- **CPU 실행**: CUDA GPU를 쓰지 않고 있다. CNN 4개 비교가 느린 핵심 요인이다. Drive 재로딩 자체에도 I/O 시간이 들지만 두 단계를 구분해야 한다.
- **과적합 신호**: 아래처럼 훈련 정확도가 약 97~99%인데 검증 정확도는 약 73~76%다. 작은 표본, 클래스 불균형, 시료별 분포 차이도 확인해야 하며 로그만으로 원인을 하나로 확정할 수 없다.
- **직전 Epoch 검증값**: 배치 진행 중에는 검증을 하지 않으므로 같은 값이 반복되는 것이 정상이다. 예를 들어 Epoch 10의 배치 로그에 표시된 검증 정확도 0.7264는 Epoch 9 결과다. Epoch 10 최종 성능은 첨부에 없다.
- **손실의 정의 차이**: 배치 Train Loss는 클래스 가중 CrossEntropy이며 Val Loss는 비가중 CrossEntropy이다. 같은 Epoch의 비가중 손실을 비교하려면 `training_history_*.csv`의 `train_loss` / `validation_loss`를 사용한다.
- **조기 종료**: `대기=5/7`은 최고 Validation Macro F1을 갱신하지 못한 Epoch가 5개 연속이라는 뜻이다. 오류 횟수가 아니다. 정확도 대신 Macro F1을 기준으로 판단하므로 검증 정확도가 오르더라도 대기가 늘 수 있다.

직전 Epoch 표시를 실제 Epoch에 맞춰 읽으면 다음과 같다.

| 완료 Epoch | 배치 훈련 정확도 | 해당 Epoch 검증 정확도 | 해당 Epoch 검증 Macro F1 |
| --- | ---: | ---: | ---: |
| 7 | 96.65% | 73.58% | 0.5740 |
| 8 | 98.03% | 76.42% | 0.6788 |
| 9 | 99.06% | 72.64% | 0.6218 |

최적 모델은 마지막 Epoch가 아니라 지금까지 가장 높은 Validation Macro F1을 기록한 체크포인트다. 더 오래 돌리는 것만으로 이 격차가 해결되지는 않는다.

## 수정 내용

`SEM_CNN_classifier_epoch3.ipynb`를 수정했다. 기존 데이터·모델 파일명을 유지하며 하이퍼파라미터, 분할 및 증강 방식은 바꾸지 않았다.

- `REQUIRE_GPU=True`: 학습 셀의 긴 Drive 검사 전에 CUDA 유무를 확인한다. 의도적 CPU 실행만 False로 허용한다. UI/기존 함수 호출은 기본적으로 CPU도 허용하면서 경고한다.
- `TRAIN_BATCH_SIZE=16`: 기존 CPU 체크포인트의 batch 16을 GPU에서도 유지한다. 배치가 바뀌면 기존 코드의 학습 설정 서명이 달라져 재개가 거부된다.
- CPU 체크포인트의 `scaler={}`를 활성 GPU GradScaler에 로드하지 않고 새 scaler를 사용한다. 모델·optimizer·scheduler·Epoch·최적 점수 등은 기존 방식대로 복원한다. GPU→GPU에서는 저장된 scaler를 복원한다. CPU와 GPU 간 수치 결과가 완전히 같음을 보장하지 않는다.
- 조기 종료 / 최대 Epoch 도달 / 실제 다음 Epoch 재개를 구분해 출력한다.
- 네 모델 중 현재 순서, 직전 검증 Epoch, 검증·저장·최종 보고서 단계를 출력한다.
- 학습 직전 안내 셀의 중복 Drive 전체 재검증을 없앴다. `train_all`의 실제 무결성 검증은 유지한다. 독립적인 증강 검증 셀도 유지한다.
- 수정된 코드 셀에 붙어 있던 예전 실행 출력은 제거했다. 사용자 Drive 체크포인트를 삭제하지 않는다.

핵심 AMP 수정은 다음과 같다. 단독으로 적용할 경우 기존 `scaler.load_state_dict(ck["scaler"])` 부분을 교체한다.

```python
scheduler.load_state_dict(ck["scheduler"])
if scaler.is_enabled() and ck.get("scaler"):
    scaler.load_state_dict(ck["scaler"])
# CPU에서 저장한 scaler={}이면 GPU의 새 GradScaler를 그대로 사용
```

## Colab에서 다음에 할 일

1. 지금 학습이 실행 중이면 중단·GPU 런타임 변경 시 **현재 Epoch의 미완료 배치**는 잃는다. 저장은 Epoch 완료 시점이다. 런타임 변경은 메모리를 초기화하지만 Drive의 파일은 남는다.
2. 수정 노트북을 열고 **런타임 → 런타임 유형 변경 → GPU**를 선택한다. GPU 할당이 실제로 되는지 `torch.cuda.is_available()`로 확인한다.
3. 설치·Drive 인증·설정·전체 Python 코드·Pipeline 생성 셀을 실행한다. **SOURCE_PATH/RESULTS_PATH는 기존 Drive 경로와 같게** 한다. 완성된 증강 데이터가 있으므로 원본 검사와 증강을 처음부터 다시 할 필요는 없다.
4. 첨부 실행의 최대 Epoch가 10이므로 설정 셀에서 `EPOCHS=10`, `REQUIRE_GPU=True`, `TRAIN_BATCH_SIZE=16`으로 둔다. 저장소 기본 `EPOCHS=3`과 다르므로 반드시 확인한다. `EPOCHS=10`은 총 10개까지라는 뜻이지 10개 추가가 아니다.
5. 장치 확인 셀과 `train_all` 셀을 실행한다. 완료된 모델은 평가하고 미완료 모델은 마지막 **완료 Epoch** 다음부터 이어간다. Early Stopping 완료 상태는 Epoch 한도를 늘려도 추가 학습하지 않는다.
6. 완료 후 metrics, confusion matrix, 클래스별 F1, 오분류 원본을 검토한다. 기존 셀에 남아 있는 출력은 과거 실행 기록이므로 현재 보고서 CSV로 판단한다. 실제 GPU 할당이나 사용자 Drive 실행은 이 환경에서 확인하지 않았다.

## 과적합을 개선하는 새 실험

먼저 현재 네 모델의 **최적 체크포인트** 비교를 끝낸다. 이후 클래스별 원본 수·독립 시료 그룹·검증 클래스별 F1·오분류·ROI·스케일바 처리부터 확인한다. Validation에 맞추어 시료를 Train으로 옮기면 평가가 편향된다.

새 학습 실험에서는 초기 LR `3e-4 → 1e-4`, weight decay `1e-4 → 1e-3`, 초기 backbone 동결 등을 각각 비교할 수 있다. 개선은 보장되지 않으며 독립 그룹을 유지한 평가가 필요하다. Train/Validation 분할·평가 기준은 동일하게 둔다.

**기존 체크포인트를 읽는 상태에서 optimizer 생성문의 LR만 바꾸면 저장된 optimizer 설정으로 덮어써진다.** 하이퍼파라미터 실험은 기존 checkpoint를 보존한 별도 모델/보고서 디렉터리와 변경 설정을 포함하는 학습 서명을 사용해야 한다. 예를 들어 기존 Pipeline의 `data`/`manifest`는 유지한 채 `models`와 `reports`를 새 빈 경로로 지정하면 기존 데이터로 새 모델을 학습할 수 있다. 이번 수정은 이런 새 실험을 자동 시작하지 않는다.

## 검증 범위

```bash
python -m unittest discover -s Fiber_SEM_Classifier/tests -v
```

10개 검사 통과: CPU/GPU 선택 가드, Epoch 재개·최대 한도·조기 종료 안내, 빈 CPU scaler 및 GPU scaler 상태 처리, 모든 Python 셀 문법. CUDA/GradScaler 분기에는 작은 대역 객체를 사용한다. 실제 PyTorch GPU 재학습·Drive 데이터 접근·정확도 향상을 검증한 결과가 아니다.
