# Case 05 — Google A3 Mega H100

**Result:** PASS · **MAPE 0.00%** · **Coverage 100% (5/5)**

Official Google documentation states A3 Mega has 8 H100 GPUs, an 8+1 physical NIC arrangement, 640 GB total HBM3, and eight GPU data networks. The shared engine receives the lower-level host profile and derives the scored quantities.

| Metric | Ref | Engine | Error |
|---|---:|---:|---:|
| GPUs/host | 8 | 8 | 0.00% |
| Physical NICs/host | 9 | 9 | 0.00% |
| GPU memory/host | 640 GB | 640 GB | 0.00% |
| GPU NIC : GPU ratio | 1.0 | 1.0 | 0.00% |
| Data networks/host | 8 | 8 | 0.00% |

The published 1,800 Gbps maximum bandwidth is stored as profile context, but is not counted as a prediction metric because it is a direct vendor specification rather than a quantity inferred by the engine.
