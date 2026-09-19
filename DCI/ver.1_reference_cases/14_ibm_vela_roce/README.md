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
