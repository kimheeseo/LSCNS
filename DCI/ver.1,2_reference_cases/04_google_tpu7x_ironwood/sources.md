# Sources — Google TPU7x / Ironwood

## Official sources

1. Google Cloud, **TPU7x (Ironwood)**  
   https://docs.cloud.google.com/tpu/docs/tpu7x

2. Google Cloud, **All Capacity mode overview**  
   https://docs.cloud.google.com/tpu/docs/all-capacity-overview

3. Google, **Ironwood: The first Google TPU for the age of inference**  
   https://blog.google/innovation-and-ai/infrastructure-and-cloud/google-cloud/ironwood-tpu-age-of-inference/

## Public facts used

- 9,216 chips per Pod.
- 4 chips per TPU7x VM/host.
- 4×4×4 cube = 64 chips and 16 hosts.
- Full Ironwood Pod = 144 cubes and 2,304 hosts.
- 3D torus topology.
- 100 Gbps DCN bandwidth per chip.
- 2,307 TFLOPs BF16 per chip.
- 4,614 TFLOPs FP8 per chip.
- 2 TensorCores and 4 SparseCores per chip.
- Google independently states 42.5 ExaFLOPS for the 9,216-chip Pod.

## Excluded

The cited sources do not establish a one-cube-to-one-physical-rack mapping. Therefore physical rack count is not inferred. Exact cable quantity, cable length, connector/patch-panel count and OCS internals are also excluded.
