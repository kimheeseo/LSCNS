# Case 02 — Google TPU v5p

**Status:** `PASS`  
**Shared engine:** `../engine/multi_arch_bom_engine.js`  
**MAPE:** **0.00%**  
**Maximum error:** **0.00%**  
**Coverage:** **100.0% (10 / 10 selected verifiable derived metrics)**  
**PASS threshold:** **< 10%**

## Method

This case uses the same shared engine as Case 01. No Google-v5p expected output is embedded in the engine.

The design input contains only architecture inputs available from Google's official documentation:

- target Pod size: 8,960 TPU chips
- 4 chips per host
- cube geometry: 4×4×4
- 1 cube per rack
- chip DCN bandwidth: 50 Gbps
- largest supported slice geometry: 16×16×24

The engine derives equipment counts and aggregate quantities.

## Reference vs engine

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| CPU hosts | 2,240 | 2,240 | 0.00% |
| Chips / cube | 64 | 64 | 0.00% |
| Cubes / Pod | 140 | 140 | 0.00% |
| Compute racks | 140 | 140 | 0.00% |
| Chips / rack | 64 | 64 | 0.00% |
| Hosts / rack | 16 | 16 | 0.00% |
| DCN bandwidth / host | 200 Gbps | 200 Gbps | 0.00% |
| Largest slice chips | 6,144 | 6,144 | 0.00% |
| Largest slice hosts | 1,536 | 1,536 | 0.00% |
| Largest slice cubes | 96 | 96 | 0.00% |

## Generic derivation examples

```
chips_per_cube = 4 × 4 × 4 = 64
host_count     = 8960 / 4 = 2240
cube_count     = ceil(8960 / 64) = 140
rack_count     = ceil(8960 / 64) = 140
hosts_per_rack = 64 / 4 = 16

host_DCN       = 50 Gbps/chip × 4 chips/host = 200 Gbps

max_slice_chips = 16 × 16 × 24 = 6144
max_slice_hosts = 6144 / 4 = 1536
max_slice_cubes = 6144 / 64 = 96
```

## Scope limitation

Google's public v5p documentation does not disclose enough detail to validate OCS count, OCS port count, exact optical cable quantity, cable length, connector/patch-panel count, or rack-level field routing. Those items are excluded rather than guessed.

The next case is Google TPU v6e / Trillium.
