# Case 09 — Meta 24,576 H100 GenAI Cluster

## 검증 목적

Meta의 전체 GPU 규모와 OCP Grand Teton의 node 내부 GPU 수를 **서로 다른 공개 Source에서 결합**해 compute-node 수를 검증한다.

- 검증 유형: **A- — Cross-source architecture derivation**
- Meta cluster: https://engineering.fb.com/2024/03/12/data-center-engineering/building-metas-genai-infrastructure/
- OCP Grand Teton: https://www.opencompute.org/documents/grand-teton-amd-based-cpu-tray-specification-v1-0-pdf

## Reference 1 — Meta cluster 규모

Meta는 다음과 같이 설명한다.

> “two versions of our 24,576-GPU data center scale cluster at Meta.”

같은 글에서:
- cluster당 24,576 H100
- 한 cluster는 RoCE (Arista 7800 + Wedge400 + Minipack2)
- 다른 cluster는 NVIDIA Quantum-2 InfiniBand
- 둘 다 400 Gbps endpoints
- 둘 다 Grand Teton 기반

을 공개한다.

## Reference 2 — Grand Teton 내부

OCP Grand Teton specification의 **Figure / Platform Block Diagram**에서 Accelerator Tray는 **GPU 0 ~ GPU 7**로 표시된다.

| Grand Teton component | 공개 구성 |
|---|---:|
| CPU Tray | 2 CPUs |
| Switch Tray | 4 PCIe Gen5 switches + 8 RDMA NICs |
| Accelerator Tray | GPU 0 … GPU 7 = **8 GPUs** |

Meta SIGCOMM 자료도 Grand Teton을 **8 GPUs + 8 RDMA NICs, 1:1 GPU:NIC** 구조로 설명한다.

## Engine derivation

```
Grand Teton nodes = 24,576 GPUs / 8 GPUs per node
                  = 3,072 nodes
```

## Reference vs Code

| Metric | Reference derivation | Engine | Error |
|---|---:|---:|---:|
| Grand Teton compute nodes | 3,072 | 3,072 | 0.00% |
| 8-GPU building blocks | 3,072 | 3,072 | 0.00% |

### Review result

- MAPE: **0.00%**
- Coverage: **2/2**
- Result: **PASS**
- Validation strength: **A-**

## 검토 결론

Case 09는 같은 문서의 숫자를 단순 합산한 것이 아니라, **Meta cluster 규모 + OCP node 내부 구성**을 교차 사용하므로 1~10 중 비교적 독립성이 높은 검증이다. 단, Meta는 전체 switch 수와 cable 수를 공개하지 않으므로 network BOM 전체 정확도를 의미하지 않는다.

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.0000% |
| Maximum error | 0.0000% |
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
