# Case 23 — NVIDIA B200 Compute Fabric

NVIDIA Table 4 gives:

| SU | Nodes | GPUs | Leaf | Spine | Compute+UFM cables | Spine-Leaf |
|---:|---:|---:|---:|---:|---:|---:|
| 4 | 127 | 1,016 | 32 | 16 | 1,020 | 1,024 |

The footnote explains the 32-node/SU design loses one DGX position for UFM connectivity.

Engine:
```
active compute links = 127×8 = 1016
UFM links = 4
Compute+UFM = 1020
fabric design slots = 128×8 = 1024
leaf = 1024/32 = 32
spine-leaf = 32×32 = 1024
spine = 1024/64 = 16
```

All four scored outputs match.

**MAPE 0.00% · PASS · Class A.**
