# Case 33 — NVIDIA DGX B300 low-density rack power hold-out

## Purpose
Frozen-engine hold-out validation for the v4.8 generalized BOM architecture model.

- Engine freeze commit: `04ad2c1e6a023811b50e79dc3b891fe5e1f590c3`
- Numerical threshold: **MAPE / per-metric error < 10%**
- Status: **PASS**
- Formal validation level: **A-**
- Semantic evidence class: **B** (per-unit power inputs are design/source inputs; the hold-out tests rack-level aggregation rather than autonomous product discovery).

## Sources
- Introduction to NVIDIA DGX B300 Systems: https://docs.nvidia.com/dgx/dgxb300-user-guide/introduction-to-dgxb300.html
- Data Center Best Practices with DGX B300: https://docs.nvidia.com/dgx-pdf/data-center-best-practices-with-dgx-b300-v1.pdf

## Reference vs engine
| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| typical_it_power_kw | 30 | 29 | 3.3333% |
| design_max_power_kw | 30 | 30 | 0.0000% |
| peak_provisioning_power_kw | 39.4 | 39.4 | 0.0000% |

**MAPE: 1.1111% · Max error: 3.3333% · PASS.**

## Interpretation
The frozen three-tier power model is applied without changing engine formulas. The NVIDIA rack-average / rack-peak values are compared with the model's typical/design/peak aggregation.
