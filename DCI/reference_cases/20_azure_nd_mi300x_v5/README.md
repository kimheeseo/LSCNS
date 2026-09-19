# Case 20 — Azure ND MI300X v5

Microsoft publishes:
| Item | Value |
|---|---:|
| MI300X GPUs | 8 |
| Memory/GPU | 192 GB |
| NICs | 8 |
| Dedicated scale-out link | 400 Gbps/GPU |
| Interconnect/VM | 3.2 Tbps |

Engine: `8×192 = 1,536 GB`; one dedicated NIC/GPU → 8 NICs.

| Metric | Reference | Engine | Error |
|---|---:|---:|---:|
| Accelerator memory | 1,536 GB | 1,536 GB | 0% |
| NIC count | 8 | 8 | 0% |

**PASS · Class C.**
