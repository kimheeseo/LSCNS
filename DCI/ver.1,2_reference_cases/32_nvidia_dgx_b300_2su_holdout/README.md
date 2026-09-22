# Case 32 — NVIDIA DGX B300 SuperPOD — 2 SU hold-out

## Purpose
Frozen-engine hold-out validation for the v4.8 generalized BOM architecture model.

- Engine freeze commit: `04ad2c1e6a023811b50e79dc3b891fe5e1f590c3`
- Numerical threshold: **MAPE / per-metric error < 10%**
- Status: **PASS**
- Formal validation level: **A**


## Sources
- DGX SuperPOD Architecture — DGX B300 / Quantum-X800 XDR: https://docs.nvidia.com/dgx-superpod/reference-architecture/scalable-infrastructure-b300-xdr/latest/dgx-superpod-architecture.html
- Major Components — DGX B300 SuperPOD: https://docs.nvidia.com/dgx-superpod/reference-architecture/scalable-infrastructure-b300-xdr/latest/components.html
- Introduction to NVIDIA DGX B300 Systems: https://docs.nvidia.com/dgx/dgxb300-user-guide/introduction-to-dgxb300.html

## Reference vs engine
| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| node_count | 144 | 144 | 0.0000% |
| accelerator_count | 1152 | 1152 | 0.0000% |
| leaf_switch_count | 16 | 16 | 0.0000% |
| spine_switch_count | 8 | 8 | 0.0000% |
| node_leaf_cable_count | 1152 | 1152 | 0.0000% |
| leaf_spine_cable_count | 1152 | 1152 | 0.0000% |

**MAPE: 0.0000% · Max error: 0.0000% · PASS.**

## Interpretation
This case was evaluated after the v1 engine was frozen. The same generic fabric-design equations are used; no expected leaf/spine/cable counts are embedded in the production calculation path.
