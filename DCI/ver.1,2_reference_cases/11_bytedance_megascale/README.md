# Case 11 — ByteDance MegaScale

## Reference evidence
USENIX NSDI 2024 paper: https://www.usenix.org/system/files/nsdi24-jiang-ziheng.pdf

The paper states: > “The total bandwidth of each Tomahawk chip is 25.6Tbps with 64×400Gbps ports.”

Its network section further specifies:

| Published architecture item | Value |
|---|---:|
| Switch layers | 3, Clos-like |
| Ports / Tomahawk-4 switch | 64 × 400G |
| Downlink : uplink | 32 : 32 |
| One 400G downlink breakout | 2 × 200G |
| NICs / server | 8 × 200G |
| Multi-rail | 8 different ToRs |
| Servers reachable by same ToR set | up to 64 |

## Code input
The engine is given a 64-server validation block, 8 server links/host, 64 logical 200G downlinks/ToR, 32 physical 400G downlinks/ToR, and 32 400G uplinks/ToR.

## Reference vs code

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Server-facing 200G links | 512 | 512 | 0.00% |
| ToRs in 8-rail group | 8 | 8 | 0.00% |
| Physical 400G downlink ports | 256 | 256 | 0.00% |
| 400G uplinks | 256 | 256 | 0.00% |

**MAPE: 0.00% · PASS · Class A-**

The paper does not disclose a complete manufacturing BOM for all 12,288 GPUs, so a full-cluster switch count is not fabricated.

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
