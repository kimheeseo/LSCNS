# Case 21 — NVIDIA DGX H100 SuperPOD

NVIDIA's official **Larger SuperPOD component counts** table gives the 4-SU row:

| SU | Nodes | GPUs | Leaf | Spine | Node-Leaf | Leaf-Spine |
|---:|---:|---:|---:|---:|---:|---:|
| 4 | 128 | 1,024 | 32 | 16 | 1,024 | 1,024 |

The common fabric engine receives 128 nodes × 8 network links and 32 down/32 up leaf allocation.

```
endpoint links = 128×8 = 1024
leaf = 1024/32 = 32
leaf-spine = 32×32 = 1024
spine = 1024/64 = 16
```

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Node→Leaf cables | 1,024 | 1,024 | 0% |
| Leaf switches | 32 | 32 | 0% |
| Leaf→Spine cables | 1,024 | 1,024 | 0% |
| Spine switches | 16 | 16 | 0% |

**MAPE 0.00% · PASS · Class A.**
