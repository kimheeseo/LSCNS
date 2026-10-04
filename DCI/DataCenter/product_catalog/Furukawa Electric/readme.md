# Furukawa Electric Data Center Catalog

Furukawa Electric's official Data Center Solutions page identifies routers, DFB laser chips, optical fiber/cable products and flexible power cables as products applicable to data centers.

## Official web sources
- Data Center Solutions: https://www.furukawaelectric.com/en/solution/datacenter/index.html
- FITELnet FX201: https://www.furukawaelectric.com/fitelnet/product/fx201/index.html
- FITELnet vFX: https://www.furukawaelectric.com/fitelnet/product/vfx/index.html
- FITELnet FX2: https://www.furukawaelectric.com/fitelnet/product/fx2/index.html
- IMDD CW-DFB laser chip: https://www.furukawaelectric.com/optical-components/en/product/signal/dfb-chip.html

## PDF sources supplied by the user
- `fttx_j417.pdf` — Furukawa Optical Cable & Products
- `em-fcc_d336.pdf` — 600V EM-FCC(T)
- `em-lmfc_e_d331.pdf` — EM-LMFC

The repository workflow `.github/workflows/furukawa-pdf-sync.yml` downloads the official Furukawa-hosted copies of those three brochures into `PDF Sources/`, so the public catalog can expose the PDF files without embedding browser-only data.

Last reviewed: 2026-10-05
