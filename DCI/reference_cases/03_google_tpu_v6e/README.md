# Case 03 — Google TPU v6e / Trillium

**Status:** `PASS`  
**Shared engine:** `../engine/multi_arch_bom_engine.js`  
**MAPE:** **0.0051%**  
**Maximum error:** **0.0460%**  
**Coverage:** **100.0% (9 / 9 selected verifiable derived metrics)**  
**PASS threshold:** **< 10%**

## Why this case matters

Case 03 is materially different from TPU v4/v5p:

- 256-chip Pod
- 8 TPU chips per host
- 2D torus rather than the v5p 3D cube hierarchy
- 4 × 200 Gbps host NICs
- explicit 25.6 Tbps Pod DCN
- 4 ICI ports per chip

The engine was extended only with generic primitives for NIC aggregation, Pod bandwidth, ICI-port aggregation and aggregate compute. No v6e expected answer was inserted.

## Reference vs engine

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Hosts / Pod | 32 | 32 | 0.0000% |
| DCN bandwidth / host | 800 Gbps | 800 Gbps | 0.0000% |
| NICs / Pod | 128 | 128 | 0.0000% |
| NIC aggregate / host | 800 Gbps | 800 Gbps | 0.0000% |
| DCN bandwidth / Pod | 25.6 Tbps | 25.6 Tbps | 0.0000% |
| ICI ports / Pod | 1,024 | 1,024 | 0.0000% |
| BF16 peak / Pod | 234.9 PFLOPs | 235.008 PFLOPs | 0.0460% |
| Full-slice chips | 256 | 256 | 0.0000% |
| Full-slice hosts | 32 | 32 | 0.0000% |

## Generic derivation examples

```
hosts = 256 / 8 = 32

NICs = 32 hosts × 4 NIC/host = 128
host_DCN = 4 × 200 Gbps = 800 Gbps
Pod_DCN = 256 chips × 100 Gbps/chip = 25.6 Tbps

ICI_ports = 256 × 4 = 1024

BF16_Pod = 918 TFLOPs/chip × 256 / 1000
         = 235.008 PFLOPs
reference = 234.9 PFLOPs (rounded official figure)

full_slice = 16 × 16 = 256 chips
slice_hosts = 256 / 8 = 32
```

The only non-zero error is the BF16 Pod value, caused by the official Pod figure being rounded to one decimal place.

## Scope limitation

No rack count is scored because the official v6e pages do not map the 256-chip Pod to a physical rack count. Exact cabling, ODF/patch-panel, connector counts and installation lengths are also not disclosed.

Next: Case 04 — Google TPU7x / Ironwood.
