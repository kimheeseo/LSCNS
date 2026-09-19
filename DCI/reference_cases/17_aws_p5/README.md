# Case 17 — AWS EC2 P5

AWS publishes:
| Field | Value |
|---|---:|
| H100 GPUs | up to 8 |
| Total HBM3 | up to 640 GB |
| EFA network | up to 3,200 Gbps |
| Local NVMe | up to 30 TB |

The engine receives 8 × 80 GB H100 and derives **640 GB** total GPU memory.

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| GPU memory / P5 | 640 GB | 640 GB | 0.00% |

**PASS · Class C.** EFA 3.2 Tbps is retained as a direct vendor specification, not a predicted output.
