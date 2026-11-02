# OFC 2026 ML Challenge | EDFA Gain Profile Prediction

**연구 목표:** 물리 기반 PI-EDFA, Spectral Attention Transformer, Residual-FiLM 구조 및 COSMOS 전이학습을 비교하여 **95채널 EDFA Gain Profile**의 예측 정확도와 미학습 운전 조건에 대한 일반화 성능을 개선합니다.

> **평가 결과 해석:** 이 저장소의 수치는 공개 데이터 기반 **내부 검증 또는 공개 TestGround 사후 평가**입니다. **Kaggle 서버 제출 이력이나 비공개 리더보드 점수를 의미하지 않습니다.** 서로 다른 데이터 분할·학습 설정에서 얻은 점수를 단일 순위로 합치지 않습니다.

## 1. 프로젝트 안내

| 연구 모듈 | GitHub 폴더 / 자료 | 주요 기능 | 현재 상태 |
|---|---|---|---|
| **01 · COSMOS Pretraining & Fine-tuning** | [COSMOS_Transfer_Learning](COSMOS_Transfer_Learning/) | COSMOS 사전학습 → 최신 Train 미세조정, PI-EDFA·Transformer의 Scratch 대조 | 내부 검증 예비 결과 확보, 전체 COSMOS·TestGround 재검증 필요 |
| **02 · Residual-FiLM Transformer** | [FiLM_Residual_Transformer](FiLM_Residual_Transformer/) | Analytical Baseline + FiLM 조건부 특징 주입 + SHB 곡률 손실 | 예비 실험 완료, SHB 오차 개선 필요 |
| **03 · Spectral Attention Transformer** | [PI_EDFA_Transformer](PI_EDFA_Transformer/) | 95채널 입력 파워·Mask·위치의 관계를 Attention으로 학습 | 공개 TestGround 사후 검증 완료 |
| **04 · Physics-Guided EDFA Benchmark** | [results_2026-11-02](results_2026-11-02/) | PI-EDFA 기준선 및 공개 데이터 검증 기록 | 비교 기준 자료 |
| **05 · PI-EDFA + FNO Residual** | [PI_EDFA_FNO_Residual](PI_EDFA_FNO_Residual/) | 물리 기반 PI-EDFA + Fourier Neural Operator 스펙트럼 잔차 보정 | 공식 TestGround 21,056개 사후 평가: **0.392948** (동일 조건 PI 0.393592) |

> 표의 01~05는 문서상의 분류입니다. **실제 폴더 이름은 변경하지 않았습니다.**

## 2. 검증 데이터와 평가 원칙

- **공식 데이터:** [OFC ML Challenge 데이터·코드 저장소](https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code)
- **입력/출력:** 95개 WDM 채널의 입력 파워, WSS 활성 Mask, 목표 Gain·Tilt 등 운전 조건 → 채널별 Gain 예측
- **학습 데이터:** `Features/Train/`의 COSMOS와 최신 Train CSV. 원본 CSV는 **Git LFS**를 통해 받아야 합니다. LFS 포인터 파일만으로는 실행되지 않습니다.
- **카테고리:** Aging / SHB (Spectral Hole Burning) / Unseen
- **평가 범위:** WSS에서 활성화된 채널만 오차 계산. 비활성 채널은 점수 대상이 아닙니다.
- **데이터 누수 방지:** 최신 Train 홀드아웃과 공개 TestGround 정답은 학습·스케일러 피팅·하이퍼파라미터 선택에 사용하지 않습니다. TestGround는 모델 설정을 확정한 뒤 **1회 사후 감사(audit)** 용도로 사용합니다.
- **제출 형식:** `Features/Test/test_features.csv`에서 예측해 `ID + 95개 채널 Gain` 컬럼을 갖는 `submission.csv`를 생성합니다.

코드에서 구현된 `official_score()`는 행별 활성 채널 MAE, 행별 절대오차 표준편차, 전체 활성 채널 P95 및 최대오차 페널티를 결합합니다. **낮을수록 좋은 점수**입니다. 공개 TestGround의 동일 수식 계산은 *Kaggle-style 로컬 재현*이지 Kaggle 서버 점수 인증은 아닙니다.

## 3. 확인된 결과 — 공개 TestGround 사후 비교 (2026-11-02)

**PI-EDFA 기준 모델 vs Spectral Transformer.** 공개 TestGround **21,056개 샘플**에 대한 별도 실험입니다.

| 모델 | Aging | SHB | Unseen | 전체 |
|---|---:|---:|---:|---:|
| PI-EDFA | **0.144064** | **0.147435** | 0.440072 | 0.382980 |
| Spectral Transformer | 0.163582 | 0.193067 | **0.364862** | **0.333409** |

**해석:** Spectral Transformer가 **Unseen 및 전체**에서 낮은 오차를 보였지만, **Aging·SHB**에서는 PI-EDFA가 우세했습니다. 이 표의 모델은 아래 **COSMOS Transfer** 실험과 학습 조건이 다르며, 해당 전이학습 모델의 TestGround 점수로 해석하면 안 됩니다.

자세한 모델 구현과 측정 기록은 [Spectral Transformer README](PI_EDFA_Transformer/README.md) 및 [Benchmark 결과](results_2026-11-02/)를 참고하세요.

## 4. COSMOS 전이학습 결과 — 최신 Train 내부 검증 (2026-11-02)

**실험 설계:** COSMOS **4,000개 표본 / 6 epoch** 사전학습, 최신 Train 학습 397개·검증 93개, Transfer/Scratch 학습률 **0.002** 통제 실험. 최적 체크포인트 기준입니다.

| 모델 | COSMOS → Train (Transfer) | Train only (Scratch) | 결과 |
|---|---:|---:|---|
| **PI-EDFA** | **0.219143** | 0.231684 | Transfer 점수 약 **5.4% 감소(개선)** |
| **Spectral Transformer** | 0.302495 | **0.274992** | Transfer 점수 약 **10.0% 증가(악화)** |

### 카테고리별 점수 (동일 93개 내부 검증 샘플)

| 모델·학습법 | Aging | SHB | Unseen | 전체 |
|---|---:|---:|---:|---:|
| PI-EDFA · Transfer | 0.143886 | 0.181089 | **0.270557** | **0.219143** |
| PI-EDFA · Scratch | **0.137763** | **0.166643** | 0.294093 | 0.231684 |
| Transformer · Transfer | 0.202078 | 0.364364 | 0.377989 | 0.302495 |
| Transformer · Scratch | **0.176621** | **0.260564** | **0.297399** | **0.274992** |

- **PI-EDFA Transfer:** Unseen에서 개선되었으나 Aging·SHB에서는 Scratch보다 악화
- **Transformer Transfer:** 세 카테고리 모두 Scratch보다 악화. COSMOS와 최신 Train 사이의 분포 차이, 스케일링 및 미세조정 전략 검토 필요
- **범위 제한:** **93개 내부 검증 표본**의 예비 실험으로, **공개 TestGround 전체 평가나 Kaggle 제출 결과가 아닙니다.**
- **상세 기록:** [COSMOS Transfer README](COSMOS_Transfer_Learning/README.md) · [동일 LR 비교 기록](COSMOS_Transfer_Learning/benchmarks/lr_control/README.md) · [비교 CSV](COSMOS_Transfer_Learning/benchmarks/lr_control/comparison.csv)

## 5. Residual-FiLM 별도 예비 실험

COSMOS **12,000개 부분집합** 사전학습 후 최신 Train을 사용한 **공개 TestGround 사후 평가**입니다. 위 COSMOS Transfer 내부 검증과 직접 비교하지 마세요.

| 학습 방식 | 전체 | Aging | SHB | Unseen |
|---|---:|---:|---:|---:|
| COSMOS → Train | 1.572768 | **1.314756** | **1.447053** | 1.608132 |
| Train only | **1.540440** | 1.342096 | 1.539642 | **1.580112** |

현재 모델의 입력 특징 복원, SHB 극단 오차 분석과 curvature loss ablation이 후속 과제입니다. [Residual-FiLM README](FiLM_Residual_Transformer/README.md)

## 6. 실행 방법 — Colab과 Kaggle-style 로컬 검증

### 주요 실행 노트북
- [COSMOS Transfer Colab](COSMOS_Transfer_Learning/COSMOS_Transfer_Colab.ipynb): COSMOS 사전학습 → 최신 Train Fine-tuning → Scratch 비교
- [COSMOS 오류 분석 Colab](COSMOS_Transfer_Learning/COSMOS_Transfer_Error_Analysis_Report.ipynb): 샘플·채널별 예측 오차와 Domain Shift 분석
- [PI-EDFA 원본 Colab](OFC_2026_Physics_Guided_EDFA_Colab.ipynb)
- [공식 측정 데이터 Benchmark Colab](OFC_2026_Official_Measured_Benchmark_Colab.ipynb)
- [공식 Subset Benchmark 스크립트](official_json_subset_benchmark.py)

### 로컬 실행

먼저 공식 데이터 저장소를 Git LFS가 설치된 환경으로 내려받고 실제 CSV가 존재하는지 확인하세요.

```bash
git lfs install
git clone https://github.com/ofc-ml-challenge/ofc-ml-challenge-data-code.git
cd ofc-ml-challenge-data-code
git lfs pull
cd ..

git clone https://github.com/kimheeseo/LSCNS.git
cd LSCNS
pip install torch numpy pandas scikit-learn matplotlib

python 26OFC_ML_challenge/COSMOS_Transfer_Learning/train_transfer.py \
  --data-root ../ofc-ml-challenge-data-code \
  --output ./ofc_cosmos_results \
  --models pi-edfa transformer \
  --pretrain-epochs 60 \
  --finetune-epochs 180
```

**사후 공개 TestGround 평가:** Train 내부 검증으로 모델·파라미터를 먼저 확정한 뒤에만 동일 명령에 `--evaluate-testground`를 추가하세요. Colab에서는 `RUN_TESTGROUND=True`로 설정할 수 있습니다. TestGround 원본의 `test_features.csv`와 `test_labels.csv` 모두 Git LFS로 받아야 합니다.

학습 완료 후 `./ofc_cosmos_results/comparison.csv` 및 `comparison.json`에 모델별 내부 검증 점수가 기록됩니다. 각 모델 하위 폴더에는 **체크포인트(.pt), 학습 이력, 검증 오류 진단, 제출용 `submission.csv`**가 생성되며, 사후 평가 옵션을 켠 경우 카테고리별 TestGround JSON도 저장됩니다.

> **상태 및 재현 주의:** README의 기존 수치는 저장된 실험 기록을 요약한 것이며, 이 문서 수정만으로 학습이나 Kaggle 제출이 새로 실행된 것은 아닙니다. 전체 COSMOS, 여러 시드, 동일 TestGround 기반의 독립 비교가 추가로 필요합니다.

## 7. PI-EDFA + FNO + Residual Learning (신규 구현)

- **모델:** 기존 PI-EDFA의 예측 Gain을 기준선으로 사용하고, **1D Fourier Neural Operator**가 채널 간 스펙트럼 의존성을 학습하여 잔차를 추가합니다.
- **학습:** 최신 Train 학습/검증 그룹 분리 → PI-EDFA 학습 → PI 동결 → FNO 잔차 학습. COSMOS를 사용한 PI 백본 사전학습은 옵션입니다.
- **비교:** 동일 홀드아웃에서 PI-only와 PI+FNO를 Aging / SHB / Unseen / 종합으로 비교하고, 모델 확정 후 공개 TestGround로 사후 평가할 수 있습니다.
- **실행:** [FNO 모델 및 평가 코드](PI_EDFA_FNO_Residual/) · [Google Colab 실행](PI_EDFA_FNO_Residual/PI_EDFA_FNO_Residual_Colab.ipynb)
- **검증 상태 (2026-11-03):** GitHub Actions에서 공식 Git LFS 데이터를 받아 공개 TestGround 전체 **21,056개** 사후 평가를 완료했습니다. 동일 60+60 Epoch 실험에서 **PI-only 0.393592 → PI+FNO 0.392948**, Aging **0.147972**, SHB **0.147722**, Unseen **0.449606**. 개선 폭은 작습니다. [실행 기록](https://github.com/kimheeseo/LSCNS/actions/runs/41540664667) · [상세 분석](PI_EDFA_FNO_Residual/README.md). **이 값은 Kaggle 비공개 리더보드 점수가 아닙니다.**

## 8. 다음 연구 과제

1. **COSMOS 전이학습 재현:** 전체 COSMOS 데이터, 여러 random seed 및 학습률·Fine-tuning 전략 비교
2. **Aging/SHB/Unseen 강건성:** 최악 채널, 전력·Tilt 조건, 활성 채널 수에 따른 오차 분석
3. **Transformer 고도화:** Residual Learning, FiLM 조건부 주입, 물리 입력 복원 및 Domain Shift 개선
4. **공정한 최종 비교:** 동일 학습 분할과 고정된 평가 프로토콜에서 Scratch/Transfer/PI-EDFA/Transformer/Residual-FiLM/FNO 비교

---
**연구 업데이트 기준:** 2026-11-03 · **기준 저장소:** [kimheeseo/LSCNS](https://github.com/kimheeseo/LSCNS) · 공식 데이터/평가 운영 권한은 OFC ML Challenge 주최 측에 있습니다.
