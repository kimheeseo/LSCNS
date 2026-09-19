# Sources — Google TPU v6e / Trillium

## Primary official sources

1. Google Cloud, **TPU v6e**  
   https://docs.cloud.google.com/tpu/docs/v6e

2. Google Cloud Compute Engine, **TPU machine specifications**  
   https://docs.cloud.google.com/compute/docs/tpus/tpu-machines

## Public facts used

- TPU Pod size: 256 chips.
- Chips per host: 8.
- Interconnect topology: 2D torus.
- ICI ports per chip: 4.
- Bidirectional ICI bandwidth per chip: 800 GB/s.
- Peak BF16 compute per chip: 918 TFLOPs.
- Official BF16 peak compute per Pod: 234.9 PFLOPs.
- Per-host NIC configuration: 4 × 200 Gbps.
- DCN bandwidth per chip: 100 Gbps.
- Data-center network bandwidth per Pod: 25.6 Tbps.
- Full Pod / largest supported slice: 16×16 = 256 chips, 32 hosts.

## Excluded

The cited public pages do not specify a physical rack count or exact rack-to-rack cable/connector/patch-panel BOM for v6e. Those quantities are not guessed and are excluded from error calculations.
