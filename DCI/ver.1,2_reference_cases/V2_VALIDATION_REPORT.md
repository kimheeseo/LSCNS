# Version 2 — 35-Case Independent Validation Report

## Outcome

- Numeric PASS (<10% each comparable metric): **35/35**
- Validation levels: **A 16 / A- 10 / B 9 / C 0**
- Highest v2 MAPE: **1.2281%** (34_nvidia_dgx_b300_4rack_power_holdout)
- Highest v2 single-metric error: **3.6842%** (Case 34 peak-provisioning power)

The numerical formulas remain shared and reference-blind. Version 2 improves the benchmark by separating numerical agreement from design independence: a low MAPE cannot by itself produce an A-grade result.

## v1 → v2 changes

1. **Input/reference separation** — `reference.json` is read only after the calculation.
2. **Independence audit** — direct count-like output fields are detected in `design_input.json`.
3. **Level gate** — A/A- requires numerical PASS, complete coverage, sufficient comparable metrics and a clean/limited input audit.
4. **Reproducibility** — each folder adds `validation_v2.json` and its README records the numeric result and evidence strength.

## Results

| # | Case | Numeric status | MAPE | Max error | Coverage | v2 level | Direct count fields |
|---:|---|---|---:|---:|---:|---|---:|
| 01 | 01_google_tpu_v4 | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |
| 02 | 02_google_tpu_v5p | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |
| 03 | 03_google_tpu_v6e | PASS | 0.0051% | 0.0460% | 100.0000% | A | 0 |
| 04 | 04_google_tpu7x_ironwood | PASS | 0.0059% | 0.0532% | 100.0000% | A | 0 |
| 05 | 05_google_a3_mega_h100 | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |
| 06 | 06_google_a3_ultra_a4 | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |
| 07 | 07_meta_rsc_phase1 | PASS | 0.0000% | 0.0000% | 100.0000% | A- | 0 |
| 08 | 08_meta_rsc_phase2 | PASS | 0.0800% | 0.1600% | 100.0000% | B | 0 |
| 09 | 09_meta_24576_h100 | PASS | 0.0000% | 0.0000% | 100.0000% | B | 0 |
| 10 | 10_meta_grand_teton_orv3 | PASS | 0.0000% | 0.0000% | 100.0000% | A- | 0 |
| 11 | 11_bytedance_megascale | PASS | 0.0000% | 0.0000% | 100.0000% | A- | 0 |
| 12 | 12_alibaba_hpn | PASS | 0.0000% | 0.0000% | 100.0000% | A- | 0 |
| 13 | 13_ibm_vela | PASS | 0.0000% | 0.0000% | 100.0000% | A- | 0 |
| 14 | 14_ibm_vela_roce | PASS | 0.0000% | 0.0000% | 100.0000% | B | 2 |
| 15 | 15_cerebras_cg1 | PASS | 0.7407% | 0.7407% | 100.0000% | B | 0 |
| 16 | 16_xai_colossus | PASS | 0.0000% | 0.0000% | 100.0000% | B | 0 |
| 17 | 17_aws_p5 | PASS | 0.0000% | 0.0000% | 100.0000% | B | 0 |
| 18 | 18_oracle_oci_h100 | PASS | 0.0000% | 0.0000% | 100.0000% | B | 0 |
| 19 | 19_azure_nd_h100_v5 | PASS | 0.0000% | 0.0000% | 100.0000% | B | 0 |
| 20 | 20_azure_nd_mi300x_v5 | PASS | 0.0000% | 0.0000% | 100.0000% | B | 0 |
| 21 | 21_nvidia_dgx_h100_superpod | PASS | 0.0000% | 0.0000% | 100.0000% | A- | 0 |
| 22 | 22_nvidia_dgx_b200_superpod | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |
| 23 | 23_nvidia_b200_compute_fabric | PASS | 0.0000% | 0.0000% | 100.0000% | A- | 0 |
| 24 | 24_nvidia_gb200_nvl72 | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |
| 25 | 25_nvidia_gb200_components | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |
| 26 | 26_nvidia_gb200_network_fabrics | PASS | 0.0000% | 0.0000% | 100.0000% | A- | 0 |
| 27 | 27_nvidia_gb300_nvl72 | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |
| 28 | 28_nvidia_dsx_ncp | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |
| 29 | 29_frontier | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |
| 30 | 30_aurora | PASS | 0.0377% | 0.2264% | 100.0000% | A | 0 |
| 31 | 31_nvidia_dgx_b300_1su | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |
| 32 | 32_nvidia_dgx_b300_2su_holdout | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |
| 33 | 33_nvidia_dgx_b300_2rack_power_holdout | PASS | 1.1111% | 3.3333% | 100.0000% | A- | 0 |
| 34 | 34_nvidia_dgx_b300_4rack_power_holdout | PASS | 1.2281% | 3.6842% | 100.0000% | A- | 0 |
| 35 | 35_nvidia_dgx_b300_18su_holdout | PASS | 0.0000% | 0.0000% | 100.0000% | A | 0 |

## Interpretation

- **A/A-**: numerical accuracy and input independence are both supported by the recorded evidence.
- **B**: numerical result passes, but architecture/policy quantities still appear in the design input.
- **C**: profile-only or evidence-limited validation; it is retained rather than overstating design autonomy.

This classification is deliberately stricter than v1. It reports where the solver should next replace reference-derived structural policies with catalog, port-packing, rack-placement and physical-link-graph solvers.


## Cases 31–35 interpretation

- **Case 31** is the new DGX B300 golden/reference case used during the v4.8 architecture update. It is not counted as a strict unseen hold-out.
- **Cases 32–35** were evaluated after the shared v1 engine freeze commit `04ad2c1e6a023811b50e79dc3b891fe5e1f590c3`.
- Cases 32–35 are **B300-family intra-family hold-outs**. They test frozen equations at additional scale/power conditions, but they are not presented as cross-vendor generalization proof.
- Case 35 excludes the published 18-SU GPU count from exact scoring because the same NVIDIA row states 1,296 DGX B300 nodes and 9,216 GPUs, which is internally inconsistent with 8 GPUs per DGX B300 system.
- A stronger next validation step is a frozen-engine **cross-vendor hold-out** using previously unseen architecture families and physical BOM fields.
