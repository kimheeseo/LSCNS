# Case 19 — Azure ND H100 v5

Microsoft's official host table lists:
| Item | Value |
|---|---:|
| H100 GPUs | 8 × 80 GB |
| NICs | 8 |
| Dedicated scale-out link | 400 Gbps/GPU |
| Interconnect / VM | 3.2 Tbps |

Engine derives **640 GB** GPU memory and **8 NICs**.

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Accelerator memory | 640 GB | 640 GB | 0% |
| NIC count | 8 | 8 | 0% |

**PASS · Class C.** The 3.2 Tbps value is direct vendor context, not a prediction metric.

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
