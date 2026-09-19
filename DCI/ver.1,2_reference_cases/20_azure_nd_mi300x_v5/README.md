# Case 20 — Azure ND MI300X v5

Microsoft publishes:
| Item | Value |
|---|---:|
| MI300X GPUs | 8 |
| Memory/GPU | 192 GB |
| NICs | 8 |
| Dedicated scale-out link | 400 Gbps/GPU |
| Interconnect/VM | 3.2 Tbps |

Engine: `8×192 = 1,536 GB`; one dedicated NIC/GPU → 8 NICs.

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Accelerator memory | 1,536 GB | 1,536 GB | 0% |
| NIC count | 8 | 8 | 0% |

**PASS · Class C.**

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
