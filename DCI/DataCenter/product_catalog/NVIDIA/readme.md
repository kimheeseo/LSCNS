# NVIDIA Data Center Product Catalog

Updated 2026-10-09.

Product groups: **Networking**, **Adapter**, **Cable**, **InfiniBand Platform**, **Networking/Switch**, **Networking/gateway**, **Data Center GPUs**, **CPU and Superchips**, **AI Factory Platforms**, and **Professional GPUs Reference**.

The web UI reads structured `catalog.json` files from `product_catalog/catalog-manifest.json`. PDF uploads alone do not enter the product list. Datasheet references contain `datasheetPath` for a repository PDF or `datasheetUrl` for an NVIDIA resource URL.

## Official product pages and datasheets

- Ethernet switching: https://www.nvidia.com/ko-kr/networking/ethernet-switching/
- BlueField platform: https://www.nvidia.com/ko-kr/networking/products/data-processing-unit/
- Ethernet SuperNIC: https://www.nvidia.com/ko-kr/networking/products/ethernet/supernic/
- ConnectX-9 SuperNIC datasheet: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/connectx-9-supernic-datasheet
- ConnectX-8 SuperNIC datasheet: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/connectx-datasheet-c
- BlueField-3 datasheet: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/datasheet-nvidia-bluefield
- ConnectX-7 NIC datasheet: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/connectx-7-datasheet
- LinkX QSA & cable accessories: https://www.nvidia.com/ko-kr/networking/ethernet/cable-accessories/
- ConnectX Ethernet adapters: https://www.nvidia.com/ko-kr/networking/ethernet-adapters/

## BOM matching policy

- `vendorReferenceOnly` keeps duplicate PDF-only entries (already represented in Networking) visible in the vendor browser but out of automatic BOM selection.
- NVIDIA LinkX QSA adapters are **Ethernet-only**, not InfiniBand.
- MC6709309 MPO-12 to eight LC connectors is a multimode SR4 splitter; do **not** automatically substitute it for single-mode DR4/FR4 or InfiniBand NDR links.
- Family-level specs are not necessarily orderable part numbers. Validate port count, product SKU, signal protocol, optics type, reach, firmware and host compatibility before procurement.
- When adding new PDF-only folders in the future, add matching structured catalog entries and update the manifest paths.

## October 9 resource library extension

Added `Networking/Technical Resources/catalog.json` for ten newly supplied NVIDIA links: Quantum-X800 overview, BlueField-3/4, ConnectX-9, Spectrum-X, SN6000, DSX Air and AI Factory/scale-in technical articles. Product datasheets appear in both their existing Networking product cards and the new technical-resource section. No reference document enters automatic BOM matching.

## NVIDIA LinkX Interconnect and Networking Software (2026-10-09)

- Product source: https://networking-docs.nvidia.com/interconnect
- `LinkX Interconnect/Current Products/catalog.json`: 41 manufacturer-indexed current-listing product families across 1600G, 800G, 400G, 200G, 100G, 25G and Accessories. The `Catalog Group` spec drives website subcategory filters.
- `LinkX Interconnect/Discontinued Products/catalog.json`: 31 entries listed under **Products No Longer For Sale** (including MCA7J60-Nxxx, whose manufacturer index points to the wrong URL). All entries are reference-only and are not BOM candidates.
- `Networking Software/catalog.json`: NVIDIA DSX Air (network/data-center simulation) and NVIDIA NetQ (network operations) with official software links. Excluded from hardware BOM.
- Existing `Adapter/catalog.json` remains canonical for MAM1Q00A-QSA and MAM1Q00A-QSA28; LinkX browsing duplicates are excluded from automatic BOM matching.
- The 400G OSFP DR4/SR4 examples explicitly specify 4×100G-PAM4 and dual InfiniBand/Ethernet protocol support, as documented in NVIDIA's model pages; do not generalize these claims to all transceivers.
- Source categories reflect the NVIDIA index on 2026-10-09 and **do not guarantee in-stock availability**. Confirm exact SKU, optical reach, cooling-shell type (IHS/RHS/finned), MPO APC/UPC, firmware and host-switch compatibility before RFQ.
- Source URLs point to individual NVIDIA documentation pages, not downloaded third-party PDF copies.

## Detailed NVIDIA NIC/DPU and Quantum Switch Specs (2026-10-09)

- Official ConnectX NIC portfolio: https://www.nvidia.com/ko-kr/networking/ethernet-adapters/
- BlueField-3: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/datasheet-nvidia-bluefield
- ConnectX-8: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/connectx-datasheet-c
- ConnectX-9: https://resources.nvidia.com/en-us-accelerated-networking-resource-library/connectx-9-supernic-datasheet
- Quantum InfiniBand switches: https://www.nvidia.com/ko-kr/networking/infiniband-switching/

Updated detailed specifications for nine existing NVIDIA Networking entries rather than duplicating those product families. Added `InfiniBand Switch Models/catalog.json` with Quantum-X800 **Q3200-RA, Q3300-LD, Q3400-RA, Q3401-RD, Q3450-LD** and Quantum-2 **QM9700, QM9790**.

Important engineering distinction: ConnectX-9 per-device total interface bandwidth is 800Gb/s, whereas a separate NVIDIA Rubin platform-level networking claim reaches 1.6Tb/s. ConnectX-8 is dual-protocol Ethernet/InfiniBand, 800Gb/s total bandwidth, with up to 400Gb/s per Ethernet port according to the selected datasheet. BlueField-3 can support InfiniBand or Ethernet at up to 400Gb/s. The CPO-based Q3450-LD front panel uses MPO12 fiber rather than pluggable transceiver modules; this compatibility fact is cataloged but the existing physical optical BOM calculation has **not** been automatically modified.

Datasheet landing-page links are preserved, and stable official PDF links are available separately in the vendor product cards where supplied. Optical reach, channelization, cooling, power supply and exact purchasable SKU must still be checked for actual deployment.
