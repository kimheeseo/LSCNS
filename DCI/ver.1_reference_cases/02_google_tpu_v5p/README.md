# Case 02 — Google TPU v5p

## 검증 목적

Google Cloud TPU v5p 공식 문서의 **Pod / host / cube / slice 구조**를 공통 엔진이 재구성하는지 검증한다.

- 검증 유형: **B — Reference-input 기반 파생 검증**
- 공식 문서: https://docs.cloud.google.com/tpu/docs/v5p

## Reference에서 실제 언급된 내용

공식 문서는 다음 문구로 전체 규모를 설명한다.

> “There are 8960 chips in a v5p Pod.”

또한 문서의 *VM, host and slice properties* 표에 다음 값이 제시된다.

| Google 표의 구분 | Cores | Chips | Hosts/VMs | Cubes |
|---|---:|---:|---:|---:|
| Host | 8 | 4 | 1 | — |
| Cube (rack) | 128 | 64 | 16 | 1 |
| Largest supported slice | 12,288 | 6,144 | 1,536 | 96 |
| v5p full Pod | 17,920 | 8,960 | 2,240 | 140 |

추가 공식 표:
- DCN bandwidth per chip = **50 Gbps**
- Interconnect topology = **3D torus**
- Maximum single-slice shape = **16×16×24 = 6,144 chips = 96 cubes**
- NIC throughput = **200 Gbps / host**

## 엔진 입력

| Input | Value |
|---|---:|
| target_accelerators | 8,960 |
| accelerators_per_host | 4 |
| building_block | 4×4×4 |
| blocks_per_rack | 1 |
| DCN/chip | 50 Gbps |
| validation_slice | 16×16×24 |

## Reference vs Code

| 검증 항목 | Reference | Engine | 계산 | Error |
|---|---:|---:|---|---:|
| Hosts / Pod | 2,240 | 2,240 | 8960/4 | 0.00% |
| Chips / cube | 64 | 64 | 4×4×4 | 0.00% |
| Cubes / Pod | 140 | 140 | 8960/64 | 0.00% |
| Compute racks | 140 | 140 | Google가 cube를 “Cube (rack)”로 표기 | 0.00% |
| Chips / rack | 64 | 64 | 4×4×4 | 0.00% |
| Hosts / rack | 16 | 16 | 64/4 | 0.00% |
| DCN / host | 200 Gbps | 200 Gbps | 50×4 | 0.00% |
| Max slice chips | 6,144 | 6,144 | 16×16×24 | 0.00% |
| Max slice hosts | 1,536 | 1,536 | 6144/4 | 0.00% |
| Max slice cubes | 96 | 96 | 6144/64 | 0.00% |

### Review result

- MAPE: **0.00%**
- Max error: **0.00%**
- Coverage: **10/10**
- Result: **PASS**

## 검토 결론

Case 02는 Google 공식 문서 자체가 cube를 명시적으로 **“Cube (rack)”**라고 정의하므로 rack count 140을 검증에 포함해도 근거가 있다. OCS 개수나 cable/connector BOM은 이 문서에서 공개되지 않아 제외한다.

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.0000% |
| Maximum error | 0.0000% |
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
