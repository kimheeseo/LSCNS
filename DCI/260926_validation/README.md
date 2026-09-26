# 260926_validation — Public Data Center BOM / Architecture Re-validation

검증일: **2026-09-26**

이 폴더는 인터넷에 공개된 데이터센터 / AI-HPC reference architecture 중 현재 BOM 엔진이 이미 지원하는 **10개 대표 사례**를 다시 실행하여, 현재 엔진의 수치 정밀도를 재확인하는 snapshot입니다.

> **중요:** 이 10개는 기존 35-case 개발/회귀 benchmark에서 architecture diversity를 기준으로 뽑은 **re-validation subset**입니다. 따라서 아래 MAPE는 **현재 엔진의 재현성 / 회귀 정밀도**를 보여주지만, 새로운 unseen hold-out generalization을 증명하는 수치는 아닙니다. Cross-vendor unseen hold-out은 별도로 유지해야 합니다.

## 파일

- `260926_validation_colab.ipynb` — Colab에서 현재 GitHub `main`을 clone하고 v2 엔진을 직접 다시 실행하는 재현용 notebook
- `validation_cases.json` — 10개 case, 공개 source URL, validation class, snapshot 기준값
- 이 `README.md` — 검증 방법, 결과, 해석

Colab:
https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/DCI/260926_validation/260926_validation_colab.ipynb

## 10개 검증 대상

| # | Public design | Vendor / org. | Validation level | Case MAPE |
|---:|---|---|---|---:|
| 1 | Google TPU v4 | Google | A | 0.000000% |
| 2 | Google TPU v6e / Trillium | Google | A | 0.005109% |
| 3 | Meta AI Research SuperCluster Phase 1 | Meta | A- | 0.000000% |
| 4 | ByteDance MegaScale network building block | ByteDance | A- | 0.000000% |
| 5 | NVIDIA DGX H100 SuperPOD | NVIDIA | A- | 0.000000% |
| 6 | NVIDIA DGX B200 SuperPOD | NVIDIA | A | 0.000000% |
| 7 | NVIDIA GB200 NVL72 rack | NVIDIA | A | 0.000000% |
| 8 | Frontier | OLCF / HPE / AMD | A | 0.000000% |
| 9 | Aurora | ALCF / HPE / Intel | A | 0.037736% |
| 10 | NVIDIA DGX B300 SuperPOD — 1 SU | NVIDIA | A | 0.000000% |

## Snapshot 결과

- PASS: **10 / 10**
- Selected comparable metrics: **62**
- Selected-field coverage: **100%**
- Case-average MAPE: **0.004284%**
- Metric-weighted MAPE: **0.004393%**
- Maximum case MAPE: **0.037736%** — Aurora
- Maximum single-metric error: **0.226415%** — Aurora aggregate endpoint injection, reference 2.12 PB/s vs engine 2.1248 PB/s
- Validation-level mix: **A 7 / A- 3**
- Threshold: **각 비교 metric error < 10%**

## 해석

이 subset에서 현재 엔진은 공개 reference와 매우 높은 수치 일치도를 보입니다. 오차가 0이 아닌 항목도 Google TPU v6e BF16 Pod peak와 Aurora aggregate injection처럼 공개값의 반올림/표기 정밀도에 민감한 항목입니다.

하지만 이 결과를 **“범용 BOM 엔진의 완전한 0.004% 오차”**로 표현하면 안 됩니다. 이유는 다음과 같습니다.

1. 이 10개는 이미 기존 35-case 개발/회귀 ledger에 포함된 사례입니다.
2. 공개 문서가 제공하지 않는 물리 cable length, exact connector SKU, patch-panel quantity, rack routing 등은 error calculation에서 제외합니다.
3. 일부 case는 architecture policy가 `design_input.json`에 이미 구조적으로 주어져 있으므로, 수치 정밀도와 autonomous architecture selection 능력은 같은 개념이 아닙니다.
4. Case 31 DGX B300 1 SU는 v4.8 architecture update에 사용된 golden case이며 strict hold-out이 아닙니다.

따라서 이 폴더가 보여주는 것은:

> **“현재 지원되는 공개 reference architecture 10개에 대해, 공개된 비교 가능 수치의 regression/reproducibility accuracy가 매우 높다.”**

입니다.

## Public source matrix

### Google TPU v4
- Jouppi et al., *TPU v4: An Optically Reconfigurable Supercomputer for Machine Learning with Hardware Support for Embeddings*
- https://arxiv.org/abs/2304.01433
- https://doi.org/10.1145/3579371.3589350

### Google TPU v6e / Trillium
- https://docs.cloud.google.com/tpu/docs/v6e
- https://docs.cloud.google.com/compute/docs/tpus/tpu-machines

### Meta RSC
- https://ai.meta.com/blog/ai-rsc/

### ByteDance MegaScale
- https://www.usenix.org/conference/nsdi24/presentation/jiang-ziheng
- https://www.usenix.org/system/files/nsdi24-jiang-ziheng.pdf

### NVIDIA DGX H100 SuperPOD
- https://docs.nvidia.com/dgx-superpod/reference-architecture-scalable-infrastructure-h100/latest/dgx-superpod-architecture.html

### NVIDIA DGX B200 SuperPOD
- https://docs.nvidia.com/dgx-superpod/reference-architecture-scalable-infrastructure-b200/latest/dgx-superpod-architecture.html

### NVIDIA GB200 NVL72
- https://docs.nvidia.com/dgx/dgxgb200-user-guide/hardware.html

### Frontier
- https://docs.olcf.ornl.gov/systems/frontier_user_guide.html

### Aurora
- https://docs.alcf.anl.gov/aurora/
- https://www.alcf.anl.gov/sites/default/files/2024-03/Aurora-Data%20Flow-28Feb24.pdf

### NVIDIA DGX B300 SuperPOD — 1 SU
- https://docs.nvidia.com/dgx-superpod/reference-architecture/scalable-infrastructure-b300-xdr/latest/dgx-superpod-architecture.html
- https://docs.nvidia.com/dgx-superpod/reference-architecture/scalable-infrastructure-b300-xdr/latest/components.html
- https://docs.nvidia.com/dgx/dgxb300-user-guide/introduction-to-dgxb300.html

## 다음 검증 단계

이 snapshot 이후의 강한 검증은 **새로운 cross-vendor Case 36–45**를 별도로 만들고, 엔진 freeze 후 reference answer를 보지 않은 상태에서 처음 계산하는 것입니다. 그때는 topology뿐 아니라 가능한 범위에서 physical cage / optic / cable / connector / rack-power 항목까지 함께 scoring하는 것이 바람직합니다.
