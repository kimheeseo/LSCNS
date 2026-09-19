# Case 27 — NVIDIA GB300 NVL72 AI Factory

## Reference evidence
NVIDIA's GB300 NVL72 logical architecture defines one rack with **18 compute trays / 72 GPUs**.

For each compute tray, the network section specifies:
- **4 × single-port ConnectX-8**, each at **800 Gb/s** for the compute fabric;
- **1 × BlueField-3 B3240 DPU** providing **2 × 400 Gb/s** converged links.

## Code input
Only the per-tray network structure and 18-tray rack size are supplied.

## Code derivation

```
CX-8 NICs/rack = 18 × 4 = 72
Compute endpoint BW = 72 × 800G = 57.6 Tbps

Converged 400G links = 18 × 2 = 36
Converged endpoint BW = 36 × 400G = 14.4 Tbps

GPUs = 18 trays × 4 GPUs = 72
```

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Compute trays | 18 | 18 | 0.00% |
| GPUs | 72 | 72 | 0.00% |
| CX-8 compute NICs | 72 | 72 | 0.00% |
| Compute endpoint bandwidth | 57.6 Tbps | 57.6 Tbps | 0.00% |
| Converged 400G links | 36 | 36 | 0.00% |
| Converged endpoint bandwidth | 14.4 Tbps | 14.4 Tbps | 0.00% |

**MAPE 0.00% · PASS · Class B+.**

Bandwidth values here are sums of published endpoint link rates; they are not claims of measured application throughput.
