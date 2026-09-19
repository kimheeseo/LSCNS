# Case 03 — Google TPU v6e / Trillium

## 검증 목적

TPU v6e 공식 표에 공개된 **Pod size, host, NIC, ICI, BF16 aggregate**를 공통 엔진으로 재계산한다.

- 검증 유형: **B — Reference-input 기반 파생 검증**
- 공식 문서: https://docs.cloud.google.com/tpu/docs/v6e

## Reference 표

Google의 *System architecture* 표에서 사용한 값만 추출하면 다음과 같다.

| Specification | Official value |
|---|---:|
| Peak BF16 compute / chip | 918 TFLOPs |
| ICI ports / chip | 4 |
| Chips / host | 8 |
| TPU Pod size | 256 chips |
| Interconnect topology | 2D torus |
| BF16 peak compute / Pod | 234.9 PFLOPs |
| NIC configuration / host | 4 × 200 Gbps |
| DCN bandwidth / Pod | 25.6 Tbps |

원문에는 다음과 같이 직접 표기되어 있다.

> “With a 256-chip footprint per Pod, v6e shares many similarities with v5e.”

## 엔진 입력

- 256 chips
- 8 chips/host
- 4 NICs/host
- 200 Gbps/NIC
- 100 Gbps DCN/chip
- 4 ICI ports/chip
- 918 TFLOPs BF16/chip
- max slice = 16×16

## Reference vs Code

| 검증 항목 | Reference | Engine | Error |
|---|---:|---:|---:|
| Hosts / Pod | 32 | 32 | 0.0000% |
| DCN / Host | 800 Gbps | 800 Gbps | 0.0000% |
| NICs / Pod | 128 | 128 | 0.0000% |
| NIC aggregate / Host | 800 Gbps | 800 Gbps | 0.0000% |
| DCN / Pod | 25.6 Tbps | 25.6 Tbps | 0.0000% |
| ICI ports / Pod | 1,024 | 1,024 | 0.0000% |
| BF16 Peak / Pod | 234.9 PFLOPs | 235.008 PFLOPs | 0.0460% |
| Full slice chips | 256 | 256 | 0.0000% |
| Full slice hosts | 32 | 32 | 0.0000% |

### 오차 원인

`918 × 256 / 1000 = 235.008 PFLOPs`이며, Google 표는 **234.9 PFLOPs**로 반올림된 aggregate 값을 제공한다.

### Review result

- MAPE: **0.0051%**
- Max error: **0.0460%**
- Coverage: **9/9**
- Result: **PASS**

## 검토 결론

rack count는 Google v6e 문서에서 physical rack mapping을 공개하지 않기 때문에 일부러 계산하지 않았다. 이 처리가 올바르며, rack/cable 값을 임의로 추정하는 것보다 검증 신뢰도가 높다.

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.0051% |
| Maximum error | 0.0460% |
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
