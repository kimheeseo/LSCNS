# LS Datacenter 3D v4.6.0 — Phase 2-A/B Browser Validation

- Date: 2026-10-11 (UTC)
- Tested code commit: `98d4b54c11aaa8cf7847c92c091089cd56c0bf17` (the subsequent commits only adjust test assertions and documentation)
- Workflow: https://github.com/kimheeseo/LSCNS/actions/runs/38105549847
- Screenshots, JSON test results and original test report: https://github.com/kimheeseo/LSCNS/actions/runs/38105549847/artifacts/11690165584
- System: GitHub Actions Ubuntu / Chromium Headless Shell / SwiftShader software WebGL
- Desktop viewport: 1440×900; simulated mobile: 390×844, DPR 2; synthetic two-touch pinch by Chrome DevTools Protocol.

## Final outcome

**62 PASS · 0 FAIL · 0 JavaScript page errors · 0 console errors**

### Phase 1 preserved

- 12-tick logarithmic FOV HUD, FOV calculations, and pointer/touch show/hide toggle.
- Wheel zoom and mobile pinch with detected canvas touch events; near-plane dynamic clipping.
- Rack inspection: door, MPO/MTP cassette and server tray extraction.
- Procedural GPU/card, package/HBM and conceptual sub-micrometre zoom stages.
- Power, cooling, core network, fiber-cut, normal scenarios and walk/worker first-person modes.
- N/N+1/2N redundancy comparisons, CSV export and existing 2D/Word guidebook links.

### Phase 2-A added and validated

- 7 campus zones × 5 work items = **35 MEP work items**.
- Real button opens checklist; multi-zone check status and aggregate progress update.
- Checklist action also adds a record to the original event log.
- Status persists through a browser reload.
- MEP CSV downloaded successfully.
- Mobile workboard stays within 390px viewport.

### Phase 2-B added and validated

- Real button opens optical connection detail panel.
- SVG diagram renders GPU/NIC→ToR/Leaf→Spine A/B→Core: copper access and dual optical trunks.
- CPO toggle updates and synchronizes with the underlying `LS3D_CONFIG`.
- Fiber-cut scenario turns the selected path red and displays existing simulated B reroute state.
- Normal recovery resets cut state.
- GPU rack inspector opens from optical panel.
- Mobile optical panel stays within viewport.

## FPS — software WebGL CI samples only

| Scene | Mean FPS | p95 frame time |
|---|---:|---:|
| Desktop campus, default | 6.5 | 183.4 ms |
| Desktop campus, low quality | 12.2 | 100 ms |
| Desktop GPU LOD | 35.1 | 49.9 ms |
| Desktop conceptual die LOD | 45.0 | 33.4 ms |
| Simulated mobile campus | 11.4 | 116.7 ms |

The values above are **not indicative of actual dedicated/integrated GPU performance**. They come from an artificial headless software rendering environment. Frame-times vary with shared CI resource conditions. Real Windows/Android/iOS Safari benchmarks remain to be done.

## Issue discovered, corrected and retested

When Phase 2 controls were appended to the existing viewport tool row, the toolbar overflowed the available scene width and mouse clicks missed the button. Phase 2 controls were moved into a small floating toolbar inside the viewport with a higher stacking context; the tested mouse/touch interactions succeeded after this fix.

A mobile-regression test erroneously expected `left`/`right` fields on Playwright's `boundingBox()`, which only exposes `x`, `y`, `width` and `height`. The test was corrected; both mobile panels fit within 390px.

## Architectural and validity constraints

- The MEP overlay consists of schematic line/box/cylinder markers, not a precision engineering model.
- Optical path annotations are conceptual and read existing links/settings; these controls do not automatically generate verified procurable SKUs or redefine BOM quantities.
- DAC/AEC default NIC→ToR is electrical; switch-side CPO is an integrated optical assembly, not an additional pluggable transceiver.
- All existing `실측 CAD/BIM이 아닌 교육·설계 검토용 3D 개념 모델` and `가상값/가정값` caveats remain applicable.

## Remaining recommended work

1. Frustum culling / geometry batching to improve the large campus view, benchmark on target hardware.
2. Real phone (Android and iPhone Safari) and Chrome/Edge/Firefox/WebKit compatibility tests.
3. Independent validation of equipment catalogue, optics/BOM and physically plausible CPO ports, connector polarity, reach and route lengths before professional deployment.
4. Optional detailed MEP BIM-import, clearance and collision checks are **not** included in the completed Phase 2 concept feature scope.
