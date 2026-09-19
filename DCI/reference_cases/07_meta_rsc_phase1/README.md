# Case 07 — Meta RSC Phase 1

## 검증 목적

Meta가 공개한 RSC Phase 1의 **DGX/GPU 규모, node endpoint bandwidth, storage tier**를 공통 엔진이 파생 계산하는지 검증한다.

- 검증 유형: **B — 공개 architecture 입력 기반 파생 검증**
- Source: https://ai.meta.com/blog/ai-rsc/

## Reference 원문

Meta는 RSC를 다음과 같이 명시한다.

> “RSC today comprises a total of 760 NVIDIA DGX A100 systems … for a total of 6,080 GPUs.”

같은 문단에서 각 DGX가 **1,600 Gb/s Quantum InfiniBand two-level Clos, no oversubscription**으로 연결된다고 설명한다.

Storage 문단의 공개값:

| Storage tier | Reference |
|---|---:|
| Pure Storage FlashArray | 175 PB |
| Penguin Altus cache | 46 PB |
| Pure Storage FlashBlade | 10 PB |
| Sum used by engine | 231 PB |

## Engine derivation

```
GPU count = 760 DGX × 8 A100/DGX = 6,080
Storage total = 175 + 46 + 10 = 231 PB
Endpoint aggregate = 760 × 1.6 Tbps = 1,216 Tbps
```

## Reference vs Code

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| GPUs | 6,080 | 6,080 | 0.00% |
| Published storage sum | 231 PB | 231 PB | 0.00% |
| Node endpoint aggregate | 1,216 Tbps | 1,216 Tbps | 0.00% |

### Review result

- MAPE: **0.00%**
- Coverage: **3/3**
- Result: **PASS**
- Validation strength: **B**

## 중요 해석

1,216 Tbps는 **760개 node endpoint 속도를 합산한 값**이며, Meta가 발표한 fabric switching capacity라고 부르면 안 된다. Switch 수와 cable 수는 원문에 공개되지 않으므로 현재 Case에서 검증하지 않는다.
