# Case 16 — xAI Colossus

NVIDIA publishes:
- current Colossus: **100,000 Hopper GPUs**
- network: Spectrum-X, SN5600, BlueField-3
- planned expansion: combined total **200,000 Hopper GPUs**

Engine check: `100,000 × 2 = 200,000`.

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Expanded GPU total | 200,000 | 200,000 | 0.00% |

**PASS, but Class C-.** Public material does not disclose a complete rack/switch/cable BOM, so this is only a scale-consistency case and should not be counted as strong switch/cable evidence.

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.0000% |
| Maximum error | 0.0000% |
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
