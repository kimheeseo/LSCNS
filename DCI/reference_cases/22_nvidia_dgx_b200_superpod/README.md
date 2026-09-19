# Case 22 — NVIDIA DGX B200 SuperPOD

The official NVIDIA table's 16-SU row is:

| SU | Nodes | GPUs | Leaf | Spine | Core | Node-Leaf | Leaf-Spine | Spine-Core |
|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| 16 | 512 | 4,096 | 128 | 128 | 64 | 4,096 | 4,096 | 4,096 |

Generic 3-tier derivation:
```
node links = 512×8 = 4096
leaf = 4096/32 = 128
leaf-spine = 128×32 = 4096
spine = 4096/32 = 128
spine-core = 128×32 = 4096
core = 4096/64 = 64
```

All six scored outputs reproduce the NVIDIA table exactly.

**MAPE 0.00% · PASS · Class A.**
