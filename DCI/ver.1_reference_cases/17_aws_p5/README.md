# Case 17 — AWS EC2 P5

AWS publishes:
| Field | Value |
|---|---:|
| H100 GPUs | up to 8 |
| Total HBM3 | up to 640 GB |
| EFA network | up to 3,200 Gbps |
| Local NVMe | up to 30 TB |

The engine receives 8 × 80 GB H100 and derives **640 GB** total GPU memory.

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| GPU memory / P5 | 640 GB | 640 GB | 0.00% |

**PASS · Class C.** EFA 3.2 Tbps is retained as a direct vendor specification, not a predicted output.

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
