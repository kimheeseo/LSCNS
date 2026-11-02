# PI-EDFA + Fourier Neural Operator + Residual Learning

**OFC 2026 ML Challenge · 95채널 EDFA Gain Profile 예측 실험**

## 연구 내용

기존 [PI-EDFA](../COSMOS_Transfer_Learning/physics_edfa.py) 기반 물리 모형은 유지하고, **1D Fourier Neural Operator(FNO)**가 입력 스펙트럼 전체의 채널 간 관계를 학습하여 PI-EDFA의 **잔차(Residual)**를 보정합니다.

```text
EDFA scalars + 95-channel powers + WSS mask
           |
           +--- PI-EDFA backbone ------ Gain_base
           |
           +--- 1D FNO spectral model - delta_Gain
                                               |
Final prediction = Gain_base + delta_Gain  <----+
```

- **1단계:** PI-EDFA backbone을 최신 Train 데이터로 학습합니다. 필요하면 COSMOS 사전학습을 추가할 수 있습니다.
- **2단계:** PI-EDFA backbone을 동결하고 FNO의 residual head만 학습합니다.
- **3단계:** **동일한 group-split 홀드아웃**에서 PI-EDFA와 PI-EDFA+FNO의 Aging / SHB / Unseen / 전체 점수를 비교합니다.
- **4단계(옵션):** 학습과 모델 선택이 끝난 뒤에만 공개 TestGround 사후 점수를 계산합니다.
- FNO는 FFT 기반 스펙트럼 혼합, local 1×1 convolution 및 전역 운전 조건 주입을 이용합니다. 잔차 출력층은 0으로 초기화해 첫 출력이 PI-EDFA와 같도록 합니다.
- **대회 1등 모델을 복제한 것은 아닙니다.** 이 코드는 독자적인 physics-guided FNO 실험 구조입니다.

## 입력 데이터

공식 저장소: [OFC ML Challenge Data](https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code)

**Git LFS로 실제 CSV**가 내려받아져야 합니다. 스크립트는 아래 경로를 사용합니다.

| 데이터 | 파일 |
|---|---|
| 최신 Train | `Features/Train/train_features.csv`, `train_labels.csv` |
| COSMOS (선택) | `Features/Train/COSMOS_features.csv`, `COSMOS_labels.csv` |
| 제출용 Test | `Features/Test/test_features.csv` |
| 사후 TestGround (선택) | `Features/TestGround/test_features.csv`, `test_labels.csv` |

## 실행

Colab에서 실행하려면 [PI-EDFA + FNO Colab](PI_EDFA_FNO_Residual_Colab.ipynb)을 사용하세요.

로컬에서:

```bash
pip install torch numpy pandas scikit-learn matplotlib
git lfs install
git clone https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code.git

python train.py --data-root /path/to/ofc-ml-challenge-data-code \
  --output ./fno_results \
  --pi-epochs 180 --fno-epochs 180
```

COSMOS PI backbone 사전학습 추가:

```bash
python train.py --data-root /path/to/ofc-ml-challenge-data-code \
  --output ./fno_cosmos_results \
  --pretrain-epochs 30 --cosmos-max-rows 4000 \
  --pi-epochs 180 --fno-epochs 180
```

- `--cosmos-max-rows 0`: COSMOS 전체 행 사용
- `--fno-width 48 --fno-modes 16 --fno-depth 3`: 기본 FNO 구조
- `--pi-epochs 2 --fno-epochs 2 --cpu`: 빠른 기능 확인, 성능 추정 금지
- `--evaluate-testground`: **모델 확정 후 사후 평가**만 실시

## 출력 파일

| 결과 파일 | 용도 |
|---|---|
| `validation_comparison.json` | 최근 Train 동일 홀드아웃 비교 (PI-only / PI+FNO, 카테고리별) |
| `testground_comparison.json` | 옵션 실행 후 공개 TestGround 사후 비교 |
| `submission.csv` | Kaggle 형식 `ID` + 95채널 Gain 예측 |
| `best_pi_edfa_fno.pt` | PI/FNO 가중치·입력 정규화 정보·모델 설정 |
| `recent_pi_history.csv` · `recent_fno_history.csv` | 학습 이력 |

## 현재 검증 수준

**완료:** PyTorch 복소 FFT 역전파, 합성 데이터 end-to-end 학습/추론, CSV 출력, TestGround 경로 기능 점검.

**미완료:** OFC 공식 측정 CSV 기반 성능 측정. 이 실행 환경에서는 공식 Git LFS 실측 CSV를 내려받지 못했으므로 **실제 공식 데이터 점수를 아직 주장하지 않습니다.** 합성 smoke-test 수치는 Kaggle 점수가 아닙니다.

기존 [COSMOS 전이학습 예비 결과](../COSMOS_Transfer_Learning/README.md)와 [PI-EDFA / Spectral Transformer 검증](../PI_EDFA_Transformer/README.md)은 **다른 실험**이므로 이 신규 FNO의 결과처럼 취급하지 마세요.
