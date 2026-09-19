# Case 12 — Alibaba HPN

## Reference evidence
SIGCOMM 2024 paper: https://cs.stanford.edu/~keithw/sigcomm2024/sigcomm24-final878-acmpaginated.pdf

The paper states: > “The two ports of each NIC are connected to different ToRs”

Key published values:

| Paper item | Value |
|---|---:|
| GPUs / host | 8 |
| Backend NICs / host | 8 |
| Ports / backend NIC | 2 × 200G |
| Segment | 1,024 active + 64 backup GPUs |
| Active hosts | 128 |
| ToRs / segment | 16 |
| Downstream / ToR | 128 active + 8 backup ×200G |
| Upstream / ToR | 60 × 400G |
| Oversubscription | 1.067:1 |

## Code derivation
```
active backend links = 128 hosts × 8 rails × 2 ToR ports = 2048
ToRs = 2048 / 128 active ports = 16
uplinks = 16 × 60 = 960
backup links = 16 × 8 = 128
```

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Active backend 200G links | 2,048 | 2,048 | 0.00% |
| ToR count | 16 | 16 | 0.00% |
| 400G uplinks | 960 | 960 | 0.00% |
| Backup 200G ports | 128 | 128 | 0.00% |

**MAPE: 0.00% · PASS · Class A**

This Case directly tests dual-ToR/rail behavior that the original single/dual/rail UI could not compose.
