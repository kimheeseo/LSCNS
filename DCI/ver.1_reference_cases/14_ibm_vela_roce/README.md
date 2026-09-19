# Case 14 — IBM Vela RoCE / Clos network

## Reference evidence
The ASPLOS paper confirms Vela's virtualized RoCE + GPU Direct architecture. IBM's architecture article supplies the deterministic Clos structure:
- redundant NIC ports terminate on different ToRs;
- each ToR connects to 4 spine switches;
- each ToR→spine relation uses 2 × 100G links.

```
TOR-spine links = 2 ToR × 4 spines × 2 links = 16
```

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| TOR-spine 100G links / redundant pair | 16 | 16 | 0.00% |

**MAPE 0.00% · PASS · Class B.**

The ASPLOS abstract's ~1500-GPU and ~80% throughput values are approximate performance statements, so they are not used as BOM MAPE references.

## Version 2 Independent Validation

| Item | v2 result |
|---|---:|
| Numerical result | PASS |
| MAPE | 0.0000% |
| Maximum error | 0.0000% |
| Coverage | 100.0000% |
| Validation level | **B** |
| Direct output-count inputs | 2 |

### Reference comparison

The engine calculates from `design_input.json` only, then compares the output with `reference.json`. The numeric error above is therefore the reference-versus-calculation error; unsupported or undisclosed fields remain outside the MAPE.

### Improvement from version 1

- Adds an explicit input-independence audit instead of treating a low numerical MAPE alone as A-grade evidence.
- Flags direct Leaf/Spine/ToR/Rack/Cable/Optic/OCS count-like fields when present in the design input.
- Exports `validation_v2.json` with MAPE, maximum error, coverage, and a validation level in one reproducible record.

### Interpretation

Structural count-like fields were provided as a design policy; the case cannot claim fully autonomous A validation until product/constraint solvers derive them.
