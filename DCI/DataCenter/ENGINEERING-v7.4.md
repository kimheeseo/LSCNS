# v7.4.0 — calculation accuracy and supported scope

This release is a conceptual engineering planning tool, not a certified purchase or construction BOM. `REVIEW` with `usable: true` provides results; `INVALID` or a non-usable result clears the BOM. Exact optical/SKU compatibility, site electrical protection, airflow/hydraulics and reference-architecture certification remain outside the solver.

## Changes

- Numeric inputs preserve zero. Empty, nonfinite, fractional count, out-of-range and unsupported values are rejected rather than silently clamped.
- Actual endpoint reservations drive logical link totals, per-device port/cage checks and segment BOM. Partial pods have balanced path multiplicity rather than an inconsistent uniform full-mesh count. Core links include their media, endpoint modules and fiber paths.
- DGX H200/B200/B300 receipt entries use the exact DGX calculation profile. Other OEMs are not substituted under DGX RU/power assumptions. Non-compute products are RFQ capacity lines rather than purported compatible catalogue matches.
- Physical module capacity uses logical ports per cage. Integrated DAC/AOC assembly lines remain RFQ until a compatible dual-port/fanout assembly is selected. Spares are procurement stock, excluded from installed loss, links, racks and power.
- OS2 channel loss exposes distance, connector pairs, splices, margin and project budget. MPO-12 uses 12 installed fibers for an 8-active-fiber channel; the provisional 800G 2×DR4 profile uses 24 installed fibers for 16 active. Smaller incompatible trunks are rejected. Trunks are packed only between the same device pair. Connector mapping/polarity/pinning and actual module loss budgets require confirmation.
- Patch/harness/adapter endpoint channel allowances and panel capacity are visible. Dedicated patch racks reserve panel RU; final route-to-rack distribution is RFQ and is not an installation schedule.
- Compute/network/storage/operations rack allocation observes reserved RU, power and cooling headroom. Single/dual ToR placement also respects downlink capacity. The representative compute rack shows the actual number and RU of servers; network equipment uses separate racks.
- PDU capacity is sized per rack/per side with load and outlet assumptions. Qualified DGX power cords are retained. DGX B200's 5+1 PSUs do not imply full-performance 2N after one feed fails.
- Storage and operations assumptions contribute power, RU, racks and PDU sizing. Their management links, storage optics, protection and throughput remain RFQ.
- Facility scenarios expose PUE and independent headroom. UPS sizes IT load; generators cover IT×PUE; transformer sizing uses power factor and kVA. Planning block capacities/utilization/redundancy are explicit.
- Cooling auto retains the stock server air path. Density alone does not approve immersion/cold-plate conversion. Liquid CDU quantities remain assumed capacity blocks; redundancy and liquid fraction require engineering.
- A custom server profile supports 400/800G, 1–8 NIC links, user GPU/node, RU, design kW, cords and cage packing. It is user-assumed, not a vendor-verified profile.
- Colocation mode generates rack/PDU/facility/cross-connect path capacity only, with no AI server/switch BOM. Tenant equipment and exact cross-connect components remain RFQ.
- Main results, actual optical quantities, facility panel, diagrams and Excel consume the design snapshot. Manual vendor/advisor exploration is explicitly reference-only. Input changes require recalculation before export.
- Excel includes requirements, results, purchase/installed/spare BOM, product evidence, engineering assumptions, all port nodes and all link routes. General nested audit summaries are bounded; graph sheets contain the complete reservations.

## Verification

210 valid/invalid engine scenarios: H200/B200/B300, IB/Ethernet, rail/single/dual/Clos, 1–4096 GPU; zero values, invalid input, partial pod, storage/operations dependencies, 24U rack, physical module counts, custom and colocation cases. Invariants cover link conservation, endpoint capacity, cage packing, integer BOM, spare separation, rack RU/power and per-rack PDU minimum. Boundary sizing at 100000 GPU is checked separately.

Full-page DOM integration checks initial render, zero spares, partial pod, custom profile, colocation isolation, invalid-result clearing, no boot script errors and 7-sheet Excel export. Source and build/test scripts are maintained in the private repository; the public repository contains the generated runtime and UI.

## Primary specification references

- DGX H200: https://docs.nvidia.com/dgx/dgxh100-user-guide/introduction-to-dgxh100.html
- DGX B200: https://docs.nvidia.com/dgx/dgxb200-user-guide/introduction-to-dgxb200.html
- DGX B300: https://docs.nvidia.com/dgx/dgxb300-user-guide/introduction-to-dgxb300.html

Supported scope is 400/800G conceptual AI fabric and colocation capacity. Other rates, DC busbars, rack-scale NVL systems, certified NVIDIA reference configurations and full tenant/storage/management connectivity need additional validated profiles.
