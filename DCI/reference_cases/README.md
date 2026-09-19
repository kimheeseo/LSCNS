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
| 02 | [Google TPU v5p](02_google_tpu_v5p/) | 3D Torus | **PASS** | **0.00%** | **100.0%** |
| 03 | [Google TPU v6e / Trillium](03_google_tpu_v6e/) | 2D Torus | **PASS** | **0.0051%** | **100.0%** |
| 04 | [Google TPU7x / Ironwood](04_google_tpu7x_ironwood/) | 3D Torus / Cube hierarchy | **PASS** | **0.0059%** | **100.0%** |
| 05 | [Google A3 Mega H100](05_google_a3_mega_h100/) | GPU + multi-NIC | **PASS** | **0.00%** | **100.0%** |
| 06 | [Google A3 Ultra / A4](06_google_a3_ultra_a4/) | GPU + multi-NIC | **PASS** | **0.00%** | **100.0%** |
| 07 | [Meta RSC Phase 1](07_meta_rsc_phase1/) | 2-level nonblocking Clos | **PASS** | **0.00%** | **100.0%** |
| 08 | [Meta RSC Phase 2](08_meta_rsc_phase2/) | Large DGX Clos | **PASS** | **0.00% exact-scored** | **100.0%** |
| 09 | [Meta 24K H100 cluster](09_meta_24576_h100/) | RoCE / Quantum-2 | **PASS** | **0.00%** | **100.0%** |
| 10 | [Meta Grand Teton / ORv3](10_meta_grand_teton_orv3/) | Rack / power | **PASS** | **0.00%** | **100.0%** |
| 11 | [ByteDance MegaScale](11_bytedance_megascale/) | 3-tier Clos + 8-rail | **PASS** | **0.00%** | **100.0%** |
| 12 | [Alibaba HPN](12_alibaba_hpn/) | Rail + Dual-ToR + Dual-Plane | **PASS** | **0.00%** | **100.0%** |
| 13 | [IBM Vela](13_ibm_vela/) | Node profile / 2-level Clos | **PASS** | **0.00%** | **100.0%** |
| 14 | [IBM Vela ASPLOS](14_ibm_vela_roce/) | Virtualized RoCE / Clos | **PASS** | **0.00%** | **100.0%** |
| 15 | [Cerebras Condor Galaxy 1](15_cerebras_cg1/) | Wafer-scale cluster | **PASS** | **0.7407%** | **100.0%** |
| 16 | [xAI Colossus](16_xai_colossus/) | Spectrum-X Ethernet | **PASS (weak)** | **0.00%** | **100.0%** |
| 17 | [AWS EC2 P5](17_aws_p5/) | EFA / GPU instance | **PASS** | **0.00%** | **100.0%** |
| 18 | [Oracle OCI H100](18_oracle_oci_h100/) | RDMA GPU cluster | **PASS** | **0.00%** | **100.0%** |
| 19 | [Azure ND H100 v5](19_azure_nd_h100_v5/) | 8-GPU / 400G fabric | **PASS** | **0.00%** | **100.0%** |
| 20 | [Azure ND MI300X v5](20_azure_nd_mi300x_v5/) | 8-GPU / 400G fabric | **PASS** | **0.00%** | **100.0%** |
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

- Case 02 Google TPU v5p: **PASS · MAPE 0.00% · Coverage 100.0% (10/10 derived metrics)**

- Case 03 Google TPU v6e / Trillium: **PASS · MAPE 0.0051% · Max error 0.0460% · Coverage 100.0% (9/9 derived metrics)**

- Case 04 Google TPU7x / Ironwood: **PASS · MAPE 0.0059% · Max error 0.0532% · Coverage 100.0% (9/9 scored metrics)**



# Review 01–10 — Source Evidence Audit

The first ten cases were re-reviewed to distinguish **reference inputs** from **engine-derived outputs** and to document the exact source evidence used in each case README.

## Validation strength

| Case | Project | Reviewed MAPE | Validation class | Interpretation |
|---:|---|---:|---|---|
| 01 | Google TPU v4 | 0.00% | B | Topology inputs → rack/link/OCS derived outputs |
| 02 | Google TPU v5p | 0.00% | B | Pod/cube/host/slice derivation |
| 03 | Google TPU v6e | 0.0051% | B | Host/NIC/ICI/Pod aggregate derivation |
| 04 | Google TPU7x | 0.0059% | B+ | Google spec + All Capacity hierarchy cross-check |
| 05 | Google A3 Mega | 0.00% | C | Vendor profile consistency; not strong predictive evidence |
| 06 | Google A3 Ultra / A4 | 0.00% | C | Vendor profile consistency across two GPU generations |
| 07 | Meta RSC Phase 1 | 0.00% | B | DGX/GPU/storage/endpoint aggregate derivation |
| 08 | Meta RSC Phase 2 | 0.00% exact-scored | B+ | 16K GPU exact check; 4.992 EFLOPS is qualitative sanity check |
| 09 | Meta 24,576 H100 | 0.00% | A- | Meta cluster scale + OCP Grand Teton node architecture |
| 10 | Meta Grand Teton / ORv3 | 0.00% | B | Rack/BBU power-BOM derivation |

### Class definitions

- **A / A-**: independent or cross-source architecture quantities are combined to predict a new count.
- **B / B+**: public reference design inputs are used to derive other published quantities.
- **C**: vendor profile/catalog consistency test. Useful for DB/BOM correctness, but weaker evidence of predictive architecture accuracy.

## Audit correction

Case 08 previously treated Meta's phrase **"almost 5 exaflops"** as exactly 5.000 EFLOPS and reported 0.16% error. This review removes that approximate prose from numerical MAPE. The engine's 4.992 EFLOPS result is retained only as a qualitative consistency check.

Case 01 now excludes direct reference inputs (4096-chip target, 136-port OCS, 8 spare ports) from accuracy scoring. Only nine derived outputs are scored.

Each Case README now contains:
1. source URL,
2. short source wording or the relevant source-table subset,
3. the exact values used as engine inputs,
4. the engine-derived quantities,
5. reference vs engine error,
6. limitations and validation-strength interpretation.

- Case 11 ByteDance MegaScale: PASS · 4/4 scored topology metrics

- Case 12 Alibaba HPN: PASS · 4/4 scored topology metrics

- Case 13 IBM Vela: PASS

- Case 14 IBM Vela RoCE: PASS

- Case 15 Cerebras CG-1: PASS · MAPE 0.7407%

- Case 16 xAI Colossus: PASS but weak C-

- Case 17 AWS P5: PASS · profile arithmetic

- Case 18 Oracle OCI H100: PASS · cross-source node derivation

- Case 19 Azure ND H100 v5: PASS · host profile consistency

- Case 20 Azure ND MI300X v5: PASS · host profile consistency
