# Case 08 — Meta RSC Phase 2

## 검증 목적

Meta Phase 2의 정확한 공개 수량과 NVIDIA A100 공식 chip 성능을 이용해 cluster scale을 cross-check한다.

- 검증 유형: **B+ — Cross-source sanity check**
- Meta: https://ai.meta.com/blog/supercomputer-meta-research-supercluster-2023/
- NVIDIA A100: https://www.nvidia.com/en-us/data-center/a100/

## Reference 원문

Meta는 다음과 같이 명시한다.

> “At full strength, we achieve almost 5 exaflops of computing power.”

그리고 동일 문단에서:
- 2,000 DGX A100 systems
- 16,000 A100 GPUs
- NVIDIA Quantum InfiniBand 16 Tb/s fabric

을 공개한다.

NVIDIA A100 공식 사양의 관련 행:

| Specification | A100 value |
|---|---:|
| BFLOAT16 Tensor Core, dense | 312 TFLOPS |
| BFLOAT16 with sparsity | 624 TFLOPS |

## Review에서 수정한 점

이전 validation은 Meta의 **“almost 5 exaflops”**를 정확히 5.000 EFLOPS라고 간주해 0.16% 오차를 계산했다. 이는 원문이 근사 표현이므로 엄밀한 MAPE 기준으로 부적절하다.

따라서 Review 이후:
- **정확한 공개값 16,000 GPUs만 scored metric**
- 4.992 EFLOPS는 **qualitative cross-check**
- “almost 5”와의 상대차 0.16%는 참고값일 뿐 MAPE에 포함하지 않음

## Engine derivation

```
GPU count = 2,000 × 8 = 16,000
BF16 dense = 16,000 × 312 TFLOPS
           = 4.992 EFLOPS
```

## Reference vs Code

| Metric | Reference | Engine | Error handling |
|---|---:|---:|---|
| A100 GPU count | 16,000 | 16,000 | **0.00% — scored** |
| Cluster BF16 | “almost 5 EFLOPS” | 4.992 EFLOPS | consistent; **not numerically scored** |
| Fabric | 16 Tb/s | not predicted | published context only |

### Review result

- Exact scored MAPE: **0.00%**
- Exact scored metrics: **1/1**
- Qualitative compute cross-check: **consistent**
- Result: **PASS**
- Validation strength: **B+**

## 검토 결론

Case 08은 오히려 이번 Review를 통해 더 엄밀해졌다. “almost” 같은 표현을 정확한 reference number로 강제 변환하지 않고, exact quantity와 qualitative sanity check를 분리한다.

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.0800% |
| Maximum error | 0.1600% |
| Coverage | 100.0000% |
| Validation level | **B** |
| Direct output-count inputs | 0 |

### Reference comparison

The engine calculates from `design_input.json` only, then compares the output with `reference.json`. The numeric error above is therefore the reference-versus-calculation error; unsupported or undisclosed fields remain outside the MAPE.

### Improvement from version 1

- Adds an explicit input-independence audit instead of treating a low numerical MAPE alone as A-grade evidence.
- Flags direct Leaf/Spine/ToR/Rack/Cable/Optic/OCS count-like fields when present in the design input.
- Exports `validation_v2.json` with MAPE, maximum error, coverage, and a validation level in one reproducible record.

### Interpretation

No direct output-count field found in design input.
