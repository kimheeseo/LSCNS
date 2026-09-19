# Case 19 — Azure ND H100 v5

Microsoft's official host table lists:
| Item | Value |
|---|---:|
| H100 GPUs | 8 × 80 GB |
| NICs | 8 |
| Dedicated scale-out link | 400 Gbps/GPU |
| Interconnect / VM | 3.2 Tbps |

Engine derives **640 GB** GPU memory and **8 NICs**.

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Accelerator memory | 640 GB | 640 GB | 0% |
| NIC count | 8 | 8 | 0% |

**PASS · Class C.** The 3.2 Tbps value is direct vendor context, not a prediction metric.
