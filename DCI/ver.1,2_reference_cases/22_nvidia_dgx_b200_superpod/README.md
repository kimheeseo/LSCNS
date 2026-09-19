# Case 22 — NVIDIA DGX B200 SuperPOD

The official NVIDIA table's 16-SU row is:

| SU | Nodes | GPUs | Leaf | Spine | Core | Node-Leaf | Leaf-Spine | Spine-Core |
|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| 16 | 512 | 4,096 | 128 | 128 | 64 | 4,096 | 4,096 | 4,096 |

Generic 3-tier derivation:
```
node links = 512×8 = 4096
leaf = 4096/32 = 128
leaf-spine = 128×32 = 4096
spine = 4096/32 = 128
spine-core = 128×32 = 4096
core = 4096/64 = 64
```

All six scored outputs reproduce the NVIDIA table exactly.

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
