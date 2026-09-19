# Case 28 — NVIDIA DSX / NCP Data Center Architecture

## Reference evidence
NVIDIA's NCP data-center architecture describes the **GB200 compute tray** as:
- 2 Grace CPUs;
- **4 B200 GPUs**;
- **CIN: 4 × 400 Gb/s ConnectX-7**;
- **TAN: 2 × 400 Gb/s BlueField-3**, each configured as **2 × 200 Gb/s**;
- SMN: 1 × 1 GbE management.

## Code derivation

```
CIN NIC count = 4
CIN endpoint bandwidth = 4 × 400G = 1.6 Tbps

TAN logical 200G links = 2 DPUs × 2 = 4
TAN endpoint bandwidth = 4 × 200G = 0.8 Tbps
```

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| B200 GPUs/tray | 4 | 4 | 0.00% |
| CIN ConnectX-7 NICs | 4 | 4 | 0.00% |
| CIN endpoint bandwidth | 1.6 Tbps | 1.6 Tbps | 0.00% |
| TAN logical 200G links | 4 | 4 | 0.00% |
| TAN endpoint bandwidth | 0.8 Tbps | 0.8 Tbps | 0.00% |

**MAPE 0.00% · PASS · Class B.**

This validates endpoint/network-BOM arithmetic for a compute tray. It does not yet validate an entire DSX data-center switch count.

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
