# Case 08 — Meta RSC Phase 2

**Result:** PASS · **MAPE 0.080%** · **Max error 0.160%** · **Coverage 100% (2/2)**

Meta reports **2,000 DGX A100 systems / 16,000 A100 GPUs**, connected by a **16 Tb/s Quantum InfiniBand fabric**, and describes the system as delivering **almost 5 exaFLOPS**.

The common engine receives 2,000 hosts, 8 GPUs/host, and NVIDIA's public dense BF16 rate of 312 TFLOPS/A100:

```
GPU count = 2000 × 8 = 16000
BF16 = 16000 × 312 TFLOPS = 4.992 EFLOPS
```

| Metric | Ref | Engine | Error |
|---|---:|---:|---:|
| A100 GPUs | 16,000 | 16,000 | 0.000% |
| BF16 cluster compute | ~5.000 EFLOPS | 4.992 EFLOPS | 0.160% |

The 16 Tb/s fabric value is retained as published context but is not treated as a prediction because Meta provides it directly and does not publish enough switch/link counts to reconstruct it from first principles.
