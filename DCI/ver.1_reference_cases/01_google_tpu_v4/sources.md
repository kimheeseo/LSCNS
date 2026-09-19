# Sources — Google TPU v4

- Primary paper: Jouppi et al., **TPU v4: An Optically Reconfigurable Supercomputer for Machine Learning with Hardware Support for Embeddings**, ISCA 2023. https://doi.org/10.1145/3579371.3589350
- Public preprint: https://arxiv.org/abs/2304.01433

## Extracted BOM facts used in validation

- 4 TPU v4 chips per CPU host.
- 64 TPU chips + 16 CPU hosts fit in one rack as a 4×4×4 block.
- 6 faces × 16 links = 96 optical links per block/rack.
- Each block connects to 48 OCSes.
- Palomar OCS: 136 ports = 128 working + 8 spare.
- 64 blocks/racks × 64 TPU = 4,096 TPU chips.
- 64 racks × 96 optical links = 6,144 rack-side optical link endpoints.
- 48 OCS × 128 working ports = 6,144 working ports.

Values not explicitly disclosed by the source (for example exact installed cable lengths, MPO/LC panel quantities, and vendor BOM part numbers) are not assigned an error percentage.
