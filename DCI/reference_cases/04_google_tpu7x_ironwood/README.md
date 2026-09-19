# Case 04 — Google TPU7x / Ironwood

**Status:** `PASS`  
**Shared engine:** `../engine/multi_arch_bom_engine.js`  
**MAPE:** **0.0059%**  
**Maximum error:** **0.0532%**  
**Coverage:** **100.0% (9 / 9 scored metrics)**  
**PASS threshold:** **< 10%**

## Independent validation point

This case includes a useful independent cross-check: Google separately states that a full 9,216-chip Ironwood Pod delivers **42.5 ExaFLOPS**. The common engine does not use that 42.5 value as input. It receives **4,614 TFLOPs FP8 per chip** and the **9,216-chip target**, then calculates:

```
4614 TFLOPs/chip × 9216 chips
= 42,522,624 TFLOPs
= 42.522624 ExaFLOPs
```

Compared with Google's rounded 42.5 ExaFLOPS, the error is **0.0532%**.

## Reference vs engine

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Hosts / Pod | 2,304 | 2,304 | 0.0000% |
| Chips / cube | 64 | 64 | 0.0000% |
| Hosts / cube | 16 | 16 | 0.0000% |
| Cubes / Pod | 144 | 144 | 0.0000% |
| DCN / host | 400 Gbps | 400 Gbps | 0.0000% |
| DCN / Pod | 921.6 Tbps | 921.6 Tbps | 0.0000% |
| FP8 peak / Pod | 42,500 PFLOPs | 42,522.624 PFLOPs | 0.0532% |
| TensorCores / Pod | 18,432 | 18,432 | 0.0000% |
| SparseCores / Pod | 36,864 | 36,864 | 0.0000% |

## Generic derivations

```
hosts = 9216 / 4 = 2304
chips_per_cube = 4 × 4 × 4 = 64
cubes = 9216 / 64 = 144
hosts_per_cube = 64 / 4 = 16

DCN_per_host = 100 × 4 = 400 Gbps
DCN_per_Pod = 9216 × 100 / 1000 = 921.6 Tbps

TensorCores = 9216 × 2 = 18432
SparseCores = 9216 × 4 = 36864
```

## Important scope rule

Unlike the previous v5p case, this validation **does not call 144 cubes “144 racks.”** The official Ironwood source describes cubes/sub-blocks, not a physical rack mapping. This prevents the BOM engine from manufacturing rack quantities that the reference does not support.

Next: Case 05 — Google A3 Mega H100.
