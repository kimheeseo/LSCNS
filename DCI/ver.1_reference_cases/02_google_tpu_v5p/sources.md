# Sources — Google TPU v5p

## Primary official sources

1. Google Cloud, **TPU v5p**  
   https://docs.cloud.google.com/tpu/docs/v5p

2. Google Cloud, **TPU machines in accelerator-optimized machine family**  
   https://docs.cloud.google.com/compute/docs/tpus/tpu-machines

## Public facts used

- Full v5p Pod: 8,960 chips.
- 4 TPU chips per host/VM.
- One cube/rack: 4×4×4 = 64 chips and 16 hosts.
- Full Pod: 2,240 hosts/VMs and 140 cubes.
- Interconnect topology: 3D torus.
- Largest supported single slice: 16×16×24 = 6,144 chips, 1,536 hosts, 96 cubes.
- DCN bandwidth: 50 Gbps per chip.
- ct5p-hightpu-4t VM network bandwidth: 200 Gbps.

Exact OCS quantity, OCS port count, installed cable count/length, connector count, and patch-panel BOM are not published on these official pages and are therefore not assigned an error percentage in this case.
