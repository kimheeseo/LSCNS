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
