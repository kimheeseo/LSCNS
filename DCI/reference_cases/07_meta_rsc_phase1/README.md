# Case 07 — Meta RSC Phase 1

**Result:** PASS · **MAPE 0.00%** · **Coverage 100% (3/3 scored derived metrics)**

Meta states that Phase 1 RSC contained **760 NVIDIA DGX A100 systems / 6,080 GPUs**, with each DGX connected by **1,600 Gb/s Quantum InfiniBand** in a **two-level non-oversubscribed Clos**, plus storage tiers of **175 PB + 46 PB + 10 PB**.

| Metric | Ref | Engine | Error |
|---|---:|---:|---:|
| A100 GPUs | 6,080 | 6,080 | 0.00% |
| Total published storage | 231 PB | 231 PB | 0.00% |
| Aggregate node-endpoint bandwidth | 1,216 Tbps | 1,216 Tbps | 0.00% |

Derivations:
- GPUs = 760 nodes × 8 GPUs/node
- Storage = 175 + 46 + 10 PB
- Endpoint bandwidth = 760 × 1.6 Tbps

Switch counts are not scored because Meta's article specifies topology class and non-oversubscription but does not publish the full switch-count BOM.
