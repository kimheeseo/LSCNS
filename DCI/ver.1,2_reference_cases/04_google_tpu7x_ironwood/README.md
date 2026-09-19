# Case 04 — Google TPU7x / Ironwood

## 검증 목적

Ironwood의 **chip-host-cube-Pod hierarchy**와 per-chip 성능을 Pod aggregate로 재계산한다.

- 검증 유형: **B+ — 공식 문서 간 cross-check 포함**
- TPU7x: https://docs.cloud.google.com/tpu/docs/tpu7x
- All Capacity: https://docs.cloud.google.com/tpu/docs/all-capacity-overview

## Reference 표

TPU7x 공식 사양에서 사용한 값:

| Specification | TPU7x |
|---|---:|
| Chips / Pod | 9,216 |
| BF16 / chip | 2,307 TFLOPs |
| FP8 / chip | 4,614 TFLOPs |
| HBM / chip | 192 GiB |
| TensorCores / chip | 2 |
| SparseCores / chip | 4 |
| DCN / chip | 100 Gbps |
| Topology | 3D torus |

All Capacity 문서의 물리 hierarchy:

| Topology concept | Chips | Hosts |
|---|---:|---:|
| Host | 4 | 1 |
| Cube | 64 | 16 |
| Full block / Pod | 9,216 = 144 cubes | 2,304 |

공식 문서는 다음과 같이 설명한다.

> “TPU7x chips have a 3D torus interconnect topology.”

## Reference vs Code

| 검증 항목 | Reference | Engine | Error |
|---|---:|---:|---:|
| Hosts / Pod | 2,304 | 2,304 | 0.0000% |
| Chips / Cube | 64 | 64 | 0.0000% |
| Hosts / Cube | 16 | 16 | 0.0000% |
| Cubes / Pod | 144 | 144 | 0.0000% |
| DCN / Host | 400 Gbps | 400 Gbps | 0.0000% |
| DCN / Pod | 921.6 Tbps | 921.6 Tbps | 0.0000% |
| FP8 Peak / Pod | 42.5 EFLOPS (Google rounded) | 42.522624 EFLOPS | 0.0532% |
| TensorCores / Pod | 18,432 | 18,432 | 0.0000% |
| SparseCores / Pod | 36,864 | 36,864 | 0.0000% |

### Review result

- MAPE: **0.0059%**
- Max error: **0.0532%**
- Coverage: **9/9**
- Result: **PASS**

## 검토 결론

Case 04는 cube를 physical rack으로 임의 해석하지 않는다. Google All Capacity 문서는 cube/sub-block hierarchy를 제공하지만 rack mapping은 제공하지 않으므로 현재 엔진이 rack 값을 N/A로 두는 것이 맞다.

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.0059% |
| Maximum error | 0.0532% |
| Coverage | 100.0000% |
| Validation level | **A** |
| Direct output-count inputs | 0 |

### Reference comparison

The engine calculates from `design_input.json` only, then compares the output with `reference.json`. The numeric error above is therefore the reference-versus-calculation error; unsupported or undisclosed fields remain outside the MAPE.

### Improvement from version 1

- Adds an explicit input-independence audit instead of treating a low numerical MAPE alone as A-grade evidence.
- Flags direct Leaf/Spine/ToR/Rack/Cable/Optic/OCS count-like fields when present in the design input.
- Exports `validation_v2.json` with MAPE, maximum error, coverage, and a validation level in one reproducible record.

### Interpretation

No direct output-count field found in design input.
