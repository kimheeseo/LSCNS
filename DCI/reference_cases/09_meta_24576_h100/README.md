# Case 09 — Meta 24,576 H100 Cluster

**Result:** PASS · **MAPE 0.00%** · **Coverage 100% (2/2 scored metrics)**

Meta publicly describes two **24,576 H100 GPU** clusters, one using a RoCE fabric based on Arista 7800 with Wedge400/Minipack2 and the other NVIDIA Quantum-2 InfiniBand; both use 400 Gbps endpoints. Meta also states both clusters are built on Grand Teton.

The OCP Grand Teton specification shows an accelerator tray with **GPU 0 through GPU 7**, i.e. 8 GPUs per integrated system.

The shared engine therefore receives only:
- target GPUs = 24,576
- GPUs per host/block = 8

and derives:

```
hosts = 24576 / 8 = 3072
Grand Teton chassis = 24576 / 8 = 3072
```

| Metric | Ref | Engine | Error |
|---|---:|---:|---:|
| Grand Teton hosts | 3,072 | 3,072 | 0.00% |
| Grand Teton 8-GPU blocks | 3,072 | 3,072 | 0.00% |

The exact number of Arista 7800, Wedge400, Minipack2, Quantum-2 switches, links, and installed cables is not disclosed in the cited Meta article, so those counts are not fabricated or scored.
