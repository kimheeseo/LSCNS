# Case 10 — Meta Grand Teton / Open Rack v3

**Result:** PASS · **MAPE 0.00%** · **Coverage 100% (3/3 scored derived metrics)**

Meta states ORv3 can support **30 kW racks**. Its improved battery backup unit provides **15 kW per shelf**, **4 minutes** of backup, and **30 kW when installed as a pair**.

The common rack-power engine receives:
- design rack power = 30 kW
- BBU shelf capacity = 15 kW
- backup duration = 4 min

and derives:

```
BBU shelves = ceil(30 / 15) = 2
pair power = 15 × 2 = 30 kW
backup duration = 4 × 60 = 240 s
```

| Metric | Ref | Engine | Error |
|---|---:|---:|---:|
| Required BBU shelves | 2 | 2 | 0.00% |
| BBU pair capacity | 30 kW | 30 kW | 0.00% |
| Backup duration | 240 s | 240 s | 0.00% |

Additional non-scored architecture facts are preserved as context: ORv3 uses 48VDC output; Grand Teton is an 8OU integrated system with CPU, switch and accelerator trays, and its OCP platform specification uses nominal ~51VDC rack input.
