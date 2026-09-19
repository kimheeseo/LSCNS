# Data Center BOM Golden Reference Validation

This directory validates the BOM/design engine in `DCI/index.html` against public, independently documented data-center and AI/HPC reference architectures.

## Method

For every case:

1. Extract only quantities explicitly supported by the primary source.
2. Store them in `reference.json`.
3. Run the current BOM engine without fitting the answer to the reference.
4. Store engine output in `calculated.json`.
5. Compare deterministic quantities in `validation.json`.
6. Report per-metric error, MAPE, coverage, and PASS/FAIL/PARTIAL/NOT_SUPPORTED.
7. Do **not** assign an error percentage to quantities the source does not disclose.

```
error_pct = abs(calculated - reference) / reference * 100
MAPE      = mean(error_pct for comparable metrics)
coverage  = comparable reference metrics / all verifiable reference metrics * 100
PASS      = error <= 10%
```

## 30-case validation matrix

| # | Reference case | Main architecture | Status | MAPE | Coverage |
|---:|---|---|---|---:|---:|
| 01 | [Google TPU v4](01_google_tpu_v4/) | 3D Torus + OCS | **PASS · RE-VALIDATED** | **0.00%** | **100.0%** |
| 02 | Google TPU v5p | 3D Torus | PENDING | — | — |
| 03 | Google TPU v6e / Trillium | 2D Torus | PENDING | — | — |
| 04 | Google TPU7x / Ironwood | TPU Pod | PENDING | — | — |
| 05 | Google A3 Mega H100 | GPU + multi-NIC | PENDING | — | — |
| 06 | Google A3 Ultra / A4 | GPU + multi-NIC | PENDING | — | — |
| 07 | Meta RSC Phase 1 | 2-level nonblocking Clos | PENDING | — | — |
| 08 | Meta RSC Phase 2 | Large DGX Clos | PENDING | — | — |
| 09 | Meta 24K H100 cluster | RoCE / Quantum-2 | PENDING | — | — |
| 10 | Meta Grand Teton / ORv3 | Rack / power | PENDING | — | — |
| 11 | ByteDance MegaScale | 3-tier Clos + 8-rail | PENDING | — | — |
| 12 | Alibaba HPN | 2-tier + Rail + Dual-ToR + Dual-Plane | PENDING | — | — |
| 13 | IBM Vela | 2-level Clos / RoCE | PENDING | — | — |
| 14 | IBM Vela ASPLOS | Virtualized RoCE | PENDING | — | — |
| 15 | Cerebras Condor Galaxy 1 | Wafer-scale cluster | PENDING | — | — |
| 16 | xAI Colossus | Spectrum-X Ethernet | PENDING | — | — |
| 17 | AWS EC2 P5 UltraCluster | EFA / GPU cluster | PENDING | — | — |
| 18 | Oracle OCI Supercluster | RDMA GPU cluster | PENDING | — | — |
| 19 | Azure ND H100 v5 | 8-GPU / 400G fabric | PENDING | — | — |
| 20 | Azure ND MI300X v5 | 8-GPU / 400G fabric | PENDING | — | — |
| 21 | NVIDIA DGX H100 SuperPOD | Leaf-Spine | PENDING | — | — |
| 22 | NVIDIA DGX B200 SuperPOD | Leaf-Spine | PENDING | — | — |
| 23 | NVIDIA B200 Compute Fabric | Leaf-Spine | PENDING | — | — |
| 24 | NVIDIA GB200 NVL72 | Rack-scale NVLink | PENDING | — | — |
| 25 | NVIDIA DGX GB200 SuperPOD Components | Rack / NIC / PSU | PENDING | — | — |
| 26 | NVIDIA GB200 Network Fabrics | NVLink / Ethernet / IB | PENDING | — | — |
| 27 | NVIDIA GB300 NVL72 AI Factory | Rack-scale fabric | PENDING | — | — |
| 28 | NVIDIA DSX/NCP DC Architecture | Rack-scale | PENDING | — | — |
| 29 | Frontier | Slingshot HPC | PENDING | — | — |
| 30 | Aurora | Slingshot HPC | PENDING | — | — |

## Current progress

- Framework/schema: **complete**
- Case 01 source extraction: **complete**
- Case 01 baseline validation: **complete**
- Shared multi-architecture engine: **created** (`engine/multi_arch_bom_engine.js`)
- Case-specific golden logic removed from `DCI/index.html`
- Case 01 re-validation via shared engine: **PASS · MAPE 0.00% · Coverage 100.0%**
- Cases 02–30: **pending**

The baseline deliberately records unsupported architectures as `NOT_SUPPORTED` instead of treating missing implementation as a fabricated 100% numerical error. Once an architecture is implemented, the same immutable reference data is used for re-test.


## Validation policy clarification

The final goal is one generalized BOM design/validation engine with <10% error across the 30 public cases. Reference answers must never be embedded in the calculation path. Each case is committed and reported independently after validation.
