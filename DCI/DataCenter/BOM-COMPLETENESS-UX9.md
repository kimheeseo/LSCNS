# DataCenter BOM completeness · 7.4.1-ux9

## Assessment and implementation

The UX8 tool covered backend topology, logical/cage capacity, distance-dependent media, structured trunks, rack limits, facility redundancy, storage equipment, operations equipment, references, and regional planning. It did not yet form a complete multi-fabric cable procurement plan.

UX9 adds:
- Basic inputs prioritize network capacity, media, and cabling. Rack power/cooling and auxiliary assumptions move to advanced inputs. NIC and switch radix overrides distinguish logical links from physical cages; supported backend rates remain 400/800G.
- Optional frontend and OOB leaf/spine fabrics include explicit port reservations, endpoint/uplink cable quantities, modules, dedicated racks, PDU/cords, switch power, and recalculated facility capacity. Storage endpoint links use the native storage switch allowance, without counting its power/switches twice; FC and Ethernet speed labels are explicit. OOB includes non-OOB device management endpoints; OOB switches' own management and storage ISLs remain a review item.
- Per-segment DAC/AEC/AOC/SR/DR/FR/LR planning policies and SMF/MMF selection. OM4 SR attenuation and connector fiber counts apply to loss and installed cores. AEC reach is a product assumption. LR requires a project optical budget. Reach alone does not guarantee compatibility or loss pass.
- Effective cable length = tray route × (1 + slack %) + 2 × per-end allowance, used consistently in media selection, channel loss and installed length. Spare stock is excluded from installed metrics.
- Required quantity = installed + ceil(installed × spare %). Purchase quantity rounds this to the selected pack. Bulk reels use floor(reel length / route length) cuts per reel, then ceil(required cables / cuts); paths cannot span reels. Termination/assembly costs remain separate.
- Liquid cooling adds planning manifold/hose allowances and optional supply/return pipe route. Server compatibility, flows, branching, fittings and CDU heat fraction remain detailed design work.
- Result summary includes total switches, installed/purchased physical modules, cable length, installed/active fibers and IT/facility power. The sensitivity tool changes one input at a time across topology, oversubscription and tray distances, reporting metric deltas rather than adding unrelated BOM units. Automated comparisons are limited to 8192 accelerators.
- Official server sources and date, explicit incomplete auxiliary optical budgets, fabric inclusion scope, procurement calculation details. Supply chain and power/optical tabs remain available.

## Limits

This is a planning BOM, not a certified installation design. Actual SKU compatibility, optical polarity/breakout, physical placement/tray coordinates, auxiliary switch RU/radix/power, storage ISLs, cable fire/region certification, mechanical cooling/hydraulics, and final quotes require project/vendor inputs. No automatic air-to-liquid conversion is inferred from rack power. No universal GPU-to-NIC rule is assumed. AOC internal fiber count is not counted as separable cable cores. Frontend/storage optical loss checks remain unverified until endpoint modules and budgets are specified. Full-system reference BOM error remains unverified; source comparisons and independent formula/route checks are shown separately.

Default H200/1024 accelerator plan with frontend/OOB enabled: 64 switches, 1676 installed optical modules, 1845 purchased modules, 152350 m cable route length, 17036 installed separable optical cores, 1920.96 kW IT power. Disabling both extra fabrics reproduces the previous backend baseline: 48 switches, 133120 m route length, 16384 installed cores.

## Validation

Run from this directory with Node and jsdom installed:

```sh
node planning-model.test.cjs
node verify-completeness.cjs
```

For a jsdom installation outside the project, set DCBOM_JSDOM_PATH to its module directory. Integration tests load actual HTML/scripts and exercise NIC/radix restoration, media/length thresholds, pack/reel rounding, auxiliary fabrics and storage no-double-counting, liquid allowances, 12 sensitivities, AMD/Intel, region PDU math, pricing, colocation, invalid-result clearing, power/optical tab content, and supply chain. JavaScript errors are asserted absent.

## Engine extension

The published engine bundle is extended with an optional per-segment media hook, per-profile installed channel cores/attenuation/budget, and 10 km input limits. No private candidate ranking database is exposed. Original hardware entries are restored after synchronous custom-switch simulations. The engine wrapper, auxiliary models and UI extension live in bom-completeness-model.js and bom-completeness.js.

Sources: NVIDIA DGX guides linked in the evidence UI; NVIDIA transceiver/fiber guide; Cisco 400G QSFP-DD module datasheet; Arista 800G FAQ. Product families are planning references, not approved SKUs.
