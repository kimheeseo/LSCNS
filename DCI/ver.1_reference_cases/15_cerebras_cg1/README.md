# Case 15 — Cerebras Condor Galaxy 1

Two official Cerebras sources are combined:
- CS-2: **850,000 AI cores/system**
- full CG-1: **64 CS-2 systems**, published as **54 million cores**

Engine: `64 × 850,000 = 54,400,000` cores.

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| AI cores | 54.0 M (rounded) | 54.4 M | 0.7407% |

**MAPE 0.7407% · PASS · Class A-.**

The small difference is consistent with the CG-1 page rounding the aggregate to “54 million”.

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.7407% |
| Maximum error | 0.7407% |
| Coverage | 100.0000% |
| Validation level | **B** |
| Direct output-count inputs | 0 |

### Reference comparison

The engine calculates from `design_input.json` only, then compares the output with `reference.json`. The numeric error above is therefore the reference-versus-calculation error; unsupported or undisclosed fields remain outside the MAPE.

### Improvement from version 1

- Adds an explicit input-independence audit instead of treating a low numerical MAPE alone as A-grade evidence.
- Flags direct Leaf/Spine/ToR/Rack/Cable/Optic/OCS count-like fields when present in the design input.
- Exports `validation_v2.json` with MAPE, maximum error, coverage, and a validation level in one reproducible record.

### Interpretation

No direct output-count field found in design input.
