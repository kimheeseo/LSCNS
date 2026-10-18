# Optical cable and connector BOM update — 2026-10-18

The main DataCenter BOM tool now distinguishes a generic optical planning profile from a selected endpoint transceiver SKU. This change addresses a concrete false pass: the old generic 400G-SR4 profile allowed 100 m and MPO-12/UPC, while NVIDIA MMA4Z00-NS400 specifies 50 m on OM4 and MPO-12/APC.

## How to use

1. In advanced design assumptions, select an optic SKU independently for Server–Leaf, Leaf–Spine and Spine–Core. `Generic planning` is deliberately an RFQ result, not a qualified product.
2. Enter the measured route distance, vertical rise, slack, connector pairs and project loss margin. The selected SKU locks its published reach, fiber type, connector and polish; an over-reach or APC/UPC conflict is an error. An incompatible 400/800G choice is also an error.
3. Keep `Trunk fiber count = Auto` to allocate one connector-compatible channel per route (12F for MPO-12, 16F for MPO-16, 2F for duplex LC). Explicit larger trunks remain possible. The audit displays active/channel/trunk fibers and warns when a trunk leaves fibers outside complete channels. Packing remains per device pair, as in the existing route graph.
4. For qualified physical 2×400G modules, enable switch-side sharing and choose one or two logical 400G ports per **server** module. Server module counts are rounded per system, never pooled across systems. Confirm the actual server/switch cage, breakout, FEC and module SKUs.
5. Check cable endpoint gender and Type-B polarity. NVIDIA's named MPO-12 module requires an unpinned female cable connector at the transceiver. An unspecified endpoint remains RFQ. A pinned male choice produces an error.

## Published optic profiles used

| Endpoint module | Reach and fiber | Connector | Published source |
|---|---|---|---|
| NVIDIA MMA4Z00-NS400 | 400G SR4, OM4, 50 m | MPO-12/APC, transceiver has pins | https://networking-docs.nvidia.com/mma4z00ns400 |
| Cisco QSFP-400G-DR4 | 400GbE DR4, OS2, 500 m | MPO-12/APC | https://www.cisco.com/c/en/us/products/collateral/interfaces-modules/400g-qsfp-transceiver-modules-ds.html |
| Cisco QSFP-400G-FR4 | 400GbE FR4, OS2, 2 km | LC duplex/UPC | https://www.cisco.com/c/en/us/products/collateral/interfaces-modules/400g-qsfp-transceiver-modules-ds.html |

These are **one-end module profiles**. Cisco examples are Ethernet modules and are flagged when selected for an InfiniBand fabric. The generic channel loss allowance in the engine is still a planning assumption, not a manufacturer-qualified optical power budget. The peer module, interoperability, route-specific polarity, cassette mapping, connector count, field test and final BOM remain project checks. `PRELIMINARY / RFQ` in the audit is intentionally not an installation approval.

## Reproducible checks

Run `node DCI/DataCenter/tests/engineering-v82.test.cjs` from the repository root. The tests cover a 60 m route that rounds to a 100 m assembly and **fails** the selected NVIDIA 50 m SKU, APC and female MPO endpoint checks, FR4 duplex 2F selection, explicit 16F/MPO-12 stranded-fiber warning, and physical twin-port module rounding. For the public 127-node Compute fabric, selected dual-port server and switch modules yield 508 server + 1,536 switch = 2,044 modules; NVIDIA lists four additional UFM modules, which this comparison deliberately excludes. The existing power, fabric and capacity tests also remain applicable.

This update does not change the previously identified 16/32-SU Core switch count discrepancy or implement a full 50-site installed-BOM benchmark. Neither item should be inferred from the optic checks.
