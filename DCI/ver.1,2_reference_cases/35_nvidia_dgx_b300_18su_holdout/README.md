# Case 35 — NVIDIA DGX B300 SuperPOD — 18 SU hold-out

## Purpose
Frozen-engine hold-out validation for the v4.8 generalized fabric / physical-cage model.

- Engine freeze commit: `04ad2c1e6a023811b50e79dc3b891fe5e1f590c3`
- Threshold: **<10%**
- Status: **PASS**
- Validation level: **A**

## Sources
- https://docs.nvidia.com/dgx-superpod/reference-architecture/scalable-infrastructure-b300-xdr/latest/dgx-superpod-architecture.html
- https://docs.nvidia.com/dgx-superpod/reference-architecture/scalable-infrastructure-b300-xdr/latest/components.html

## Reference vs engine
| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Nodes | 1,296 | 1,296 | 0.0000% |
| GPUs | 10,368 | 10,368 | 0.0000% |
| Leaf | 144 | 144 | 0.0000% |
| Spine | 72 | 72 | 0.0000% |
| Node–Leaf | 10,368 | 10,368 | 0.0000% |
| Leaf–Spine | 10,368 | 10,368 | 0.0000% |

**MAPE: 0.0000% · Max error: 0.0000% · PASS.**

The engine was not modified after the freeze for this case. It uses the same generic fabric equations and the same logical-link-to-physical-cage rules used by the earlier B300 cases.
