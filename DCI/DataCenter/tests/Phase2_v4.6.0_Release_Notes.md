# LS Datacenter 3D v4.6.0 — Phase 2 Release Notes

## Scope

Phase 2-A and 2-B are implemented as additive JavaScript/CSS modules. The Phase 1 camera, FOV HUD, semantic zoom, GPU procedural geometry and the original WebGL renderer remain in the original files. No additional runtime libraries were introduced.

- Main page: `DCI/DataCenter/LS_Datacenter_3D.html`
- Phase 2-A: `LS_Datacenter_3D_Phase2A.js`
- Phase 2-B: `LS_Datacenter_3D_Phase2B.js`
- UI styles: `LS_Datacenter_3D_Phase2.css`
- Previous modules: `LS_Datacenter_3D_Extensions.js`, `LS_Datacenter_3D_Extensions.css`

## Phase 2-A: Revit-style MEP work log

A floating `▤ MEP 작업` button opens a responsive field-engineering workboard. There are **7 areas × 5 checklist items = 35 items** covering rooftop drainage, louvres, pipework, MEP service clearance, cable trays, containment, hangers, support columns, rack cooling and optical patching.

- Checklist status persists to browser localStorage `lsdc-phase2-mep-v1`.
- Every checklist change is appended to the workboard log **and** the original event log through `__LS3D_TEST__.addEventLog()`.
- Checklists can be exported as a UTF-8 BOM-prefixed CSV.
- A few completed selections add small procedural conceptual pipe/tray/support overlays in the active zone. They are **not** construction-level 3D models.
- Checked status means reviewed/illustrated, **not built, measured, certified, or commissioned**.

## Phase 2-B: Optical and copper connection journey

A floating `◎ 광 연결` button opens the GPU NIC → ToR/Leaf → Spine A/B → Core network view, with clickable controls for the optical model and failure scenario.

- The server NIC → ToR access link is modeled as DAC/AEC **copper** by default, with **no** pluggable optical transceivers on that section.
- ToR → Spine uses a conceptual SMF/MPO optical trunk.
- Pluggable mode uses switch-side pluggable transceiver representation. CPO mode represents integrated switch-side optics and MPO, not an additional switch-side pluggable optical module.
- CPO mode is synchronized with the existing `optMode` control in the advanced optical topology model.
- A/B path coloring, cut indicator, bypass state and latency/bandwidth assumptions are read from the existing `LS3D_CONFIG` and current `fiber-cut` scenario.
- The GPU card/server tray/package phase-1 LOD renderer receives small, distance-limited electrical/optical path conceptual accents. GPU/HBM internal package is **not a fiber-optic connection**.
- Optical transceiver model numbers, measured fiber loss, detailed CPO thermal/power packaging, true path lengths and purchasing SKUs are **not** implied.

## Preserved behavior

Scenario selection (power/cooling/network/fiber cut), redundancy N/N+1/2N, worker follow/first-person, walk mode, rack inspector and mechanical animation, CSV export, events, trend graphs, original 2D link and Word guidebook remain in the original HTML/Extensions implementation. Phase 2 adds controls, it does not remove these features.

## Test / performance limitations

Automated coverage: `.github/workflows/ls-datacenter-3d-phase2-browser.yml`, based on the prior 46-check Phase 1 browser suite, and extended for MEP checks, persistence, event log, optical model toggle, cut/bypass/restoration, rack inspection and desktop/mobile panel layout.

CI uses **headless Chromium and SwiftShader software WebGL on Ubuntu**; measured FPS is not a benchmark for NVIDIA/AMD/Apple discrete/integrated GPU hardware. Physical iPhone/Android, Safari/WebKit, Windows and long-duration load testing should be conducted before external demonstration.

Website: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?v=4.6.0
