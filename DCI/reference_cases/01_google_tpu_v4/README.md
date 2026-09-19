# Case 01 — Google TPU v4

**Status:** `NOT_SUPPORTED` (baseline)  
**Tool:** `DCI/index.html` baseline 2026-09-19  
**MAPE:** N/A  
**Coverage:** 0.0%  
**Pass criterion:** error ≤ 10% for every comparable deterministic BOM metric.

## Reference architecture

Google TPU v4 uses a 4×4×4 electrical building block per rack and optical circuit switches (OCSes) to form a reconfigurable 3D torus across 64 racks.

| Metric | Reference | Current tool | Error | Status |
|---|---:|---:|---:|---|
| TPU chips | 4,096 | — | — | NOT_SUPPORTED |
| CPU hosts | 1,024 | — | — | NOT_SUPPORTED |
| Compute racks | 64 | — | — | NOT_SUPPORTED |
| TPU / rack | 64 | — | — | NOT_SUPPORTED |
| Optical links / rack | 96 | — | — | NOT_SUPPORTED |
| Rack→OCS link endpoints | 6,144 | — | — | NOT_SUPPORTED |
| OCS count | 48 | — | — | NOT_SUPPORTED |
| OCS ports / unit | 136 | — | — | NOT_SUPPORTED |
| Working ports / OCS | 128 | — | — | NOT_SUPPORTED |
| Spare ports / OCS | 8 | — | — | NOT_SUPPORTED |
| Working OCS ports total | 6,144 | — | — | NOT_SUPPORTED |
| Spare OCS ports total | 384 | — | — | NOT_SUPPORTED |

## Baseline conclusion

The current DCI BOM engine cannot be mapped to this case without changing the model. It supports leaf/spine-style `single`, `dual`, and `rail` modes but does not contain TPU-v4, 3D-torus, or OCS primitives. Therefore the baseline result is recorded as **NOT_SUPPORTED rather than assigning an artificial 100% error**.

## Required implementation before re-test

1. TPU/host allocation profile.
2. `torus3d` base topology.
3. OCS device with total/working/spare ports.
4. Rack↔OCS bidirectional optical-link model.
5. Golden validation runner that compares generated quantities with `reference.json`.

After these features are implemented this case will be re-run, and only genuinely calculated values will be used for MAPE.
