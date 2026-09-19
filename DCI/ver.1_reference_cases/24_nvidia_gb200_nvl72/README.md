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

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.0000% |
| Maximum error | 0.0000% |
| Coverage | 100.0000% |
| Validation level | **A** |
| Direct output-count inputs | 0 |

### Reference comparison

The engine calculates from `design_input.json` only, then compares the output with `reference.json`. The numeric error above is therefore the reference-versus-calculation error; unsupported or undisclosed fields remain outside the MAPE.

### Improvement from version 1

- Adds an explicit input-independence audit instead of treating a low numerical MAPE alone as A-grade evidence.
- Flags direct Leaf/Spine/ToR/Rack/Cable/Optic/OCS count-like fields when present in the design input.
- Exports `validation_v2.json` with MAPE, maximum error, coverage, and a validation level in one reproducible record.

### Interpretation

No direct output-count field found in design input.
