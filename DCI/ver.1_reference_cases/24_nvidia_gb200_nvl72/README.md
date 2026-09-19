# Case 24 — NVIDIA GB200 NVL72 rack

NVIDIA hardware guide lists an NVL72 rack with:
- **18 × 1RU compute trays**
- each tray: **2 Grace CPUs + 4 Blackwell GPUs**
- **9 × NVLink switch trays**
- each switch tray: **2 NVSwitch chips**
- **2 ToR switches** for management

Derived totals:
```
GPU = 18×4 = 72
CPU = 18×2 = 36
NVSwitch = 9×2 = 18
```

| Metric | Ref | Engine | Error |
|---|---:|---:|---:|
| Compute trays | 18 | 18 | 0% |
| GPUs | 72 | 72 | 0% |
| Grace CPUs | 36 | 36 | 0% |
| NVLink switch trays | 9 | 9 | 0% |
| NVSwitch chips | 18 | 18 | 0% |
| Mgmt ToRs | 2 | 2 | 0% |

**MAPE 0.00% · PASS · Class A.**
