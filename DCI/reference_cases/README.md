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
| 21 | [NVIDIA DGX H100 SuperPOD](21_nvidia_dgx_h100_superpod/) | Rail Leaf-Spine | **PASS** | **0.00%** | **100.0%** |
| 22 | [NVIDIA DGX B200 SuperPOD](22_nvidia_dgx_b200_superpod/) | 3-tier Leaf-Spine-Core | **PASS** | **0.00%** | **100.0%** |
| 23 | [NVIDIA B200 Compute Fabric](23_nvidia_b200_compute_fabric/) | Detailed Leaf-Spine | **PASS** | **0.00%** | **100.0%** |
| 24 | [NVIDIA GB200 NVL72](24_nvidia_gb200_nvl72/) | Rack-scale NVLink | **PASS** | **0.00%** | **100.0%** |
| 25 | [NVIDIA GB200 Components](25_nvidia_gb200_components/) | Rack / storage / PSU | **PASS** | **0.00%** | **100.0%** |
| 26 | [NVIDIA GB200 Network Fabrics](26_nvidia_gb200_network_fabrics/) | SLG Leaf-Spine | **PASS** | **0.00%** | **100.0%** |
| 27 | [NVIDIA GB300 NVL72 AI Factory](27_nvidia_gb300_nvl72/) | Rack-scale fabric | **PASS** | **0.00%** | **100.0%** |
| 28 | [NVIDIA DSX/NCP DC Architecture](28_nvidia_dsx_ncp/) | Rack-scale | **PASS** | **0.00%** | **100.0%** |
| 29 | [Frontier](29_frontier/) | Slingshot HPC | **PASS** | **0.00%** | **100.0%** |
| 30 | [Aurora](30_aurora/) | Slingshot HPC | **PASS** | **0.0377%** | **100.0%** |

## Current progress

- Framework/schema: **complete**
- Case 01 source extraction: **complete**
- Case 01 baseline validation: **complete**
- Shared multi-architecture engine: **created** (`engine/multi_arch_bom_engine.js`)
- Case-specific golden logic removed from `DCI/index.html`
- Case 01 re-validation via shared engine: **PASS · MAPE 0.00% · Coverage 100.0%**
- Cases 02–30: **complete**

All 30 benchmark folders now contain source evidence, design inputs, calculated outputs, validation metrics, and README analysis. Unsupported or undisclosed fields remain excluded rather than assigned fabricated errors.


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

- Case 21 NVIDIA H100 SuperPOD: PASS · official BOM-table reconstruction

- Case 22 NVIDIA B200 SuperPOD: PASS · 6/6 3-tier BOM metrics

- Case 23 NVIDIA B200 detailed compute fabric: PASS

- Case 24 NVIDIA GB200 NVL72 rack: PASS · 6/6 rack components

- Case 25 NVIDIA GB200 components: PASS · 7/7 rack BOM metrics

- Case 26 NVIDIA GB200 fabric: PASS · 4/4 cross-document metrics


# Review 11–30 — Source Evidence Audit

Cases 11–30 were reviewed using the same rule applied to Cases 01–10: **direct reference inputs are separated from engine-derived outputs**, and each folder README identifies the exact paper/vendor table or wording used as evidence.

| Case | Project | Reviewed MAPE | Validation class | Main evidence used |
|---:|---|---:|---|---|
| 11 | ByteDance MegaScale | 0.00% | A- | NSDI'24 Tomahawk-4 64×400G, 32/32 split, 400→2×200G, 8-rail locality |
| 12 | Alibaba HPN | 0.00% | A | SIGCOMM'24 8 GPUs/host, dual-ToR, 128+8 down, 60 up, 16 ToRs/segment |
| 13 | IBM Vela | 0.00% | C | IBM node table: 8×A100, 2 CPUs, 1.5TB DRAM, 4×3.2TB NVMe |
| 14 | IBM Vela ASPLOS/RoCE | 0.00% exact-profile | C | A100 deployment profile; approximate performance values not forced into exact MAPE |
| 15 | Cerebras CG-1 | 0.3704% | B+ | 64 CS-2 × 850k cores vs rounded 54M-core cluster figure |
| 16 | xAI Colossus | 0.00% | C- | NVIDIA 100k Hopper → stated 200k expansion; physical BOM undisclosed |
| 17 | AWS EC2 P5 | 0.00% | B+ | p5.48xlarge 8 H100 + 3.2Tbps EFA; 20k-GPU UltraCluster |
| 18 | Oracle OCI H100 | 0.00% | B+ | BM.GPU.H100.8 + 16,384-GPU Supercluster maximum |
| 19 | Azure ND H100 v5 | 0.00% | C+ | 8 H100 + 8×400G dedicated IB = 3.2Tbps |
| 20 | Azure ND MI300X v5 | 0.00% | C+ | 8 MI300X + 8×400G dedicated IB = 3.2Tbps |
| 21 | NVIDIA DGX H100 SuperPOD | 0.00% | A | NVIDIA official compute-fabric component/cable table |
| 22 | NVIDIA DGX B200 SuperPOD | 0.00% | A | NVIDIA official larger SuperPOD Leaf/Spine/Core and cable table |
| 23 | NVIDIA B200 Compute Fabric | 0.00% | A | 31/63/95/127-node detailed compute-fabric table |
| 24 | NVIDIA GB200 NVL72 | 0.00% | A- | 18 compute trays, 72 GPUs, 9 NVLink switch trays, 2 management ToRs |
| 25 | NVIDIA GB200 Components | 0.00% | A- | compute tray, power shelf, PSU and NVMe component counts |
| 26 | NVIDIA GB200 Network Fabrics | 0.00% | B+ | 4×CX-7 and 2×BF3 per tray; rack-level endpoint aggregation |
| 27 | NVIDIA GB300 NVL72 | 0.00% | B+ | 18 trays, 4×800G CX-8/tray, 2×400G converged links/tray |
| 28 | NVIDIA DSX/NCP | 0.00% | B | GB200 tray: 4×400G CIN + BF3 TAN logical links |
| 29 | Frontier | 0.00% | A- | OLCF rack/node/MI250X/GCD/NIC hierarchy |
| 30 | Aurora | 0.0377% | A | ALCF node/GPU/NIC hierarchy + independent 2.12PB/s system injection cross-check |

## Important review notes

- **Case 11** validates the disclosed 64-server network building block only; the paper does not publish a full 12,288-GPU switch BOM.
- **Case 14** keeps “~1500 GPUs / ~80% / ~70%” as approximate context rather than exact numerical truth.
- **Case 16** is intentionally marked weak because the NVIDIA announcement gives scale and network family but not enough node/switch/cable detail for a strong BOM validation.
- **Case 29** has a source-version discrepancy: the current Frontier User Guide says 9,856 compute nodes, while another OLCF allocation page contains 9,408. The Case freezes the User Guide as benchmark source.
- **Case 30** independently reconstructs 2.1248 PB/s from node × NIC × 200G; the official presentation rounds this to 2.12 PB/s, producing the 0.2264% max error.

## 30-case completion state

- Benchmark cases created: **30 / 30**
- Cases with numeric PASS (<10%): **30 / 30**
- Highest reviewed MAPE among the 30 cases: **0.3704% (Cerebras CG-1)**
- Highest single-metric reviewed error: **0.7407% (Cerebras core-count cross-check)**
- Strongest architecture/BOM cases for future regression: **Alibaba HPN, NVIDIA H100/B200 SuperPOD, NVIDIA GB200/GB300, Frontier, Aurora**
- Weaker profile/announcement cases retained but explicitly labeled: **Google A3 profiles, IBM profile checks, xAI announcement**
