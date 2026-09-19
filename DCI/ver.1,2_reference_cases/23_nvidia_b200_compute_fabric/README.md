# Case 23 — NVIDIA B200 Compute Fabric

NVIDIA Table 4 gives:

| SU | Nodes | GPUs | Leaf | Spine | Compute+UFM cables | Spine-Leaf |
|---:|---:|---:|---:|---:|---:|---:|
| 4 | 127 | 1,016 | 32 | 16 | 1,020 | 1,024 |

The footnote explains the 32-node/SU design loses one DGX position for UFM connectivity.

Engine:
```
active compute links = 127×8 = 1016
UFM links = 4
Compute+UFM = 1020
fabric design slots = 128×8 = 1024
leaf = 1024/32 = 32
spine-leaf = 32×32 = 1024
spine = 1024/64 = 16
```

All four scored outputs match.

**MAPE 0.00% · PASS · Class A.**

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.0000% |
| Maximum error | 0.0000% |
| Coverage | 100.0000% |
| Validation level | **A-** |
| Direct output-count inputs | 0 |

### Reference comparison

The engine calculates from `design_input.json` only, then compares the output with `reference.json`. The numeric error above is therefore the reference-versus-calculation error; unsupported or undisclosed fields remain outside the MAPE.

### Improvement from version 1

- Adds an explicit input-independence audit instead of treating a low numerical MAPE alone as A-grade evidence.
- Flags direct Leaf/Spine/ToR/Rack/Cable/Optic/OCS count-like fields when present in the design input.
- Exports `validation_v2.json` with MAPE, maximum error, coverage, and a validation level in one reproducible record.

### Interpretation

No direct output-count field found in design input.
