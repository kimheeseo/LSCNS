# Case 18 — Oracle OCI H100 Supercluster

Oracle publishes **16,384 H100 GPUs** as an OCI Supercluster scale point. Its H100 shape document independently defines `BM.GPU.H100.8` as **8 H100 GPUs/node** with 3.2 Tb/s RDMA.

```
nodes = 16,384 / 8 = 2,048
```

| Metric | Cross-source reference | Engine | Error |
|---|---:|---:|---:|
| H100 nodes at max H100 scale | 2,048 | 2,048 | 0.00% |

**PASS · Class A-.** Physical rack/switch counts are not published and are not guessed.

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
