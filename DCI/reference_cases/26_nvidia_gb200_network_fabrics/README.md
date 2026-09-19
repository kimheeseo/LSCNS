# Case 26 — NVIDIA GB200 Network Fabrics

## Reference evidence
NVIDIA's GB200 network-fabric guide defines **one Scale Unit (SU) as 8 DGX GB200 systems/racks = 576 GPUs**.

The same guide describes **4 Scalable Leaf Groups (SLGs)** per SU, each containing:
- 8 leaf switches
- 6 spine switches

The GB200 hardware guide independently gives 18 compute trays per rack and 4 GPUs per tray.

## Code derivation

```
compute trays = 8 racks × 18 = 144
GPUs = 144 trays × 4 = 576

leaf = 4 SLG × 8 = 32
spine = 4 SLG × 6 = 24
```

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Compute trays / SU | 144 | 144 | 0.00% |
| GPUs / SU | 576 | 576 | 0.00% |
| Leaf switches / SU | 32 | 32 | 0.00% |
| Spine switches / SU | 24 | 24 | 0.00% |

**MAPE 0.00% · PASS · Class A-.**

This is a cross-document reconstruction: rack internals come from the hardware guide, while SU/SLG network counts come from the reference architecture.
