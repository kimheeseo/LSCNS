# Case 30 — ALCF Aurora

## 검증 목적
Aurora 공식 시스템 사양의 **rack / node / GPU / Slingshot endpoint** 구조와 전체 peak injection bandwidth를 공통 HPC 엔진으로 검증한다.

- Validation class: **A — HPC architecture + independent aggregate cross-check**
- Sources:
  - https://docs.alcf.anl.gov/aurora/
  - https://www.alcf.anl.gov/sites/default/files/2024-03/Aurora-Data%20Flow-28Feb24.pdf

## Reference 표

| Published item | Official value |
|---|---:|
| Compute racks | 166 |
| Compute nodes | 10,624 |
| GPUs | 63,744 |
| GPUs / node | 6 |
| Slingshot NICs / node | 8 |
| NIC link rate | 200 Gb/s |
| Peak injection bandwidth | 2.12 PB/s |
| Topology | Dragonfly |

Aurora Machine Overview는 “10,624-node” 시스템이며 node당 6 GPU와 8 Slingshot NIC를 명시한다.

## Engine derivation
```
GPUs = 10,624 × 6 = 63,744
NICs = 10,624 × 8 = 84,992
Injection/node = 8 × 200G = 1.6 Tbps
Nodes/rack = 10,624 / 166 = 64

Aggregate injection
= 10,624 × 8 × 200 Gb/s
= 16.9984 Pb/s
= 2.1248 PB/s
```

## Reference vs Engine

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Nodes | 10,624 | 10,624 | 0.000% |
| GPUs | 63,744 | 63,744 | 0.000% |
| NICs | 84,992 | 84,992 | 0.000% |
| Injection / node | 1.6 Tbps | 1.6 Tbps | 0.000% |
| Nodes / rack | 64 | 64 | 0.000% |
| Peak aggregate injection | 2.12 PB/s | 2.1248 PB/s | **0.2264%** |

**MAPE: 0.0377% · Max error: 0.2264% · PASS.**

## 오차 원인
ALCF presentation의 2.12 PB/s는 소수 둘째 자리로 반올림된 시스템 수치다. Node×NIC×link-rate로 재구성하면 2.1248 PB/s가 되므로 약 0.2264% 차이가 난다.

## 검토 결론
Case 30은 per-node input에서 system-wide 공식 aggregate bandwidth를 독립적으로 재구성하므로 30개 중 강한 A-class validation이다.

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.0377% |
| Maximum error | 0.2264% |
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
