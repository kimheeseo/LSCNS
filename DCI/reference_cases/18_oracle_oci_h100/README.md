# Case 18 — Oracle OCI H100 Supercluster

Oracle publishes **16,384 H100 GPUs** as an OCI Supercluster scale point. Its H100 shape document independently defines `BM.GPU.H100.8` as **8 H100 GPUs/node** with 3.2 Tb/s RDMA.

```
nodes = 16,384 / 8 = 2,048
```

| Metric | Cross-source reference | Engine | Error |
|---|---:|---:|---:|
| H100 nodes at max H100 scale | 2,048 | 2,048 | 0.00% |

**PASS · Class A-.** Physical rack/switch counts are not published and are not guessed.
