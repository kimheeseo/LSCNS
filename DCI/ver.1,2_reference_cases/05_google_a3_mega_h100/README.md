# Case 05 — Google A3 Mega H100

## 검증 목적

Google A3 Mega 공식 machine table과 accelerator-networking 표를 기준으로 **GPU / physical NIC / GPU memory / data network** 구성이 공통 profile 계산과 일치하는지 확인한다.

- 검증 유형: **C — Profile consistency validation**
- 의미: 공개된 machine profile을 엔진에 넣어 BOM 필드가 일관되게 생성되는지 확인한다.
- 주의: 이 Case의 0%는 topology 예측력이 아니라 **profile ingestion / BOM consistency**를 의미한다.

## Reference 1 — Compute Engine machine table

https://docs.cloud.google.com/compute/docs/gpus

| Machine type | Physical NIC | Max network | GPU count | GPU memory |
|---|---:|---:|---:|---:|
| a3-megagpu-8g | 9 | 1,800 Gbps | 8 | 640 GB HBM3 |

## Reference 2 — Accelerator networking table

https://docs.cloud.google.com/kubernetes-engine/docs/how-to/config-auto-net-for-accelerators

| Machine | GPUs | Titanium NICs | GPU NICs | GPUDirect | Additional VPC |
|---|---:|---:|---:|---|---:|
| A3 Mega | 8 H100 | 1 | 8 | TCPXO | 8 |

Google의 생성 가이드에는 다음 문구도 있다.

> “A3 Mega requires eight data networks.”

## Reference vs Code

| 검증 항목 | Reference | Engine | Error |
|---|---:|---:|---:|
| GPUs / host | 8 | 8 | 0.00% |
| Physical NICs / host | 9 | 9 | 0.00% |
| GPU memory / host | 640 GB | 640 GB | 0.00% |
| GPU NIC : GPU ratio | 8:8 = 1 | 1 | 0.00% |
| Data networks / host | 8 | 8 | 0.00% |

### Review result

- MAPE: **0.00%**
- Coverage: **5/5**
- Result: **PASS**
- Validation strength: **C / consistency**

## 검토 결론

A3 Mega는 공개 machine table이 매우 구체적이므로 장비 profile DB 검증에는 좋은 Case다. 그러나 GPU 수, NIC 수, memory가 이미 vendor profile 입력으로 주어지므로 이것만으로 “설계 툴이 장비 수를 예측했다”고 표현하면 안 된다. 향후 cluster-size 입력에서 node/rack/switch/cable을 계산하는 별도 Case와 함께 사용해야 한다.

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
