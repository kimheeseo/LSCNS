# Planning summary and input upgrade — 2026-10-04

- Result top: installed route length, active/installed fibers, IT/facility power, BOM line count and optional quotation subtotal.
- Length includes route length × installed trunk count or integrated/P2P cable paths. Spare stock, unknown patch-lead lengths, installation slack and integrated AOC internal fibers are excluded. Colocation without route distances says uncalculated.
- Reference table compares NVIDIA DGX H200/B200/B300 node count/RU/design power to manufacturer user guides. The current B300 runtime uses 15 kW versus 14.5 kW in the official guide; the table exposes the 3.448% difference rather than silently declaring zero error. Separate rows check optical-loss arithmetic and reserved-link conservation. These checks do not establish full reference-topology, measured, certification or construction accuracy; that error remains explicitly unverified.
- Basic inputs remain visible. Port/optical/rack/facility/regional assumptions move into advanced inputs. Keyboard-accessible expandable help explains terms. Custom accelerator assumptions stay visible in basic mode.
- AMD MI300X and Intel Gaudi 3 use editable custom nodes. Manufacturer accelerator specifications provide context; OEM server RU, total power, external 400G link count and cords are planning assumptions. The generic Ethernet solver does not reproduce Gaudi-native scale-out wiring. NVIDIA switch capacity used by this generic solver remains an RFQ choice.
- Region presets are selectable example supplies, not universal country requirements. Voltage/current/phases/power factor optionally size PDU kW, then the existing engine sizes PDU quantities. Frequency and IEC/UL/site procurement review selections are metadata, not certification or automatic product compatibility checks.
- Quote unit prices are entered by the user and stored separately per currency and item. Zero is allowed; missing or negative prices are excluded. Partial pricing never becomes a complete estimate. Taxes, transport and construction are excluded. No currency conversion occurs.
- Snapshot fields `planningTotals`, `validation`, `region`, `planningQuote`, and extended inputs pass through existing Excel audit export.

## Validation

Run `node planning-model.test.cjs` beside the runtime. It verifies the three official node references, independent known cable/fiber totals, spare exclusion, quote currency separation, intentional reference mismatch detection, single/three-phase PDU sizing and colocation isolation.

DOM integration was checked for initial summary, custom preset edits, AMD/Intel vendor identity, regional PDU counts, partial quote editing, colocation, invalid-result clearing, optical/cooling tab rendering and restored Supply Chain. No script exceptions were observed. A full graphical browser was not available in the authoring runtime.

## Sources

- H200: https://docs.nvidia.com/dgx/dgxh100-user-guide/introduction-to-dgxh100.html
- B200: https://docs.nvidia.com/dgx/dgxb200-user-guide/introduction-to-dgxb200.html
- B300: https://docs.nvidia.com/dgx/dgxb300-user-guide/introduction-to-dgxb300.html
- MI300X: https://www.amd.com/en/products/accelerators/instinct/mi300/mi300x.html
- Gaudi 3: https://www.intel.com/content/www/us/en/content-details/845118/intel-gaudi-3-ai-accelerator-30-3-30-pdf.html
