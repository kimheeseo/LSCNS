# Case 25 — NVIDIA DGX GB200 SuperPOD Components

## Reference evidence
NVIDIA's component guide defines the rack-level component building blocks used here.

| Published component | Reference value |
|---|---:|
| Compute trays / rack | 18 |
| Data NVMe / compute tray | 4 × 3.84 TB |
| Boot NVMe / compute tray | 1 × 1.92 TB |
| Power shelves / rack | 8 |
| PSUs / power shelf | 6 |
| PSU rating | 5.5 kW |

The engine does not receive the rack totals below; it receives the per-tray/per-shelf component profile.

## Code derivation

```
PSU count = 8 shelves × 6 = 48
Installed PSU nameplate = 48 × 5.5 = 264 kW

Data NVMe count = 18 trays × 4 = 72
Data NVMe raw capacity = 72 × 3.84 = 276.48 TB

Boot NVMe count = 18 × 1 = 18
Boot NVMe raw capacity = 18 × 1.92 = 34.56 TB
```

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Power shelves | 8 | 8 | 0.00% |
| PSUs | 48 | 48 | 0.00% |
| Installed PSU capacity | 264 kW | 264 kW | 0.00% |
| Data NVMe count | 72 | 72 | 0.00% |
| Data NVMe raw capacity | 276.48 TB | 276.48 TB | 0.00% |
| Boot NVMe count | 18 | 18 | 0.00% |
| Boot NVMe raw capacity | 34.56 TB | 34.56 TB | 0.00% |

**MAPE 0.00% · PASS · Class A.**

The 264 kW figure is installed PSU nameplate capacity, not a claim that the rack continuously consumes 264 kW.
