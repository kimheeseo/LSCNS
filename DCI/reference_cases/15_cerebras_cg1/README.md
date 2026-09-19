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
