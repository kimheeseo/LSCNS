# Changelog

## v3.2.0 — 2026-11-28
- Aligned both BMH marine sleeves and reinforcement rings to the exact shared cable curve; no route/control-point or power-equation changes.
- Default Auto quality with conservative desktop/touch/CPU starting tiers and a Medium ceiling for automatic upgrades. Medium DPR cap 1, simpler water mesh, no Medium/Low or underwater real-time shadows.
- Distance culling for reef/fish/details and distant underwater land facilities; idle redraw suppression with correct CPU-throttle retry and completed-render frame timing.
- 59/59 r128 geometry/function regression checks passed. Both sleeve centerlines remain within 0.003 scene units of the cable route. Browser CPU-fallback inspection confirmed BMH interior/shore connection and controls; GPU/mobile FPS unverified. Scene-budget comparison and scope are in tests/performance-20261128.json.

## v3.0 — 2026-11-27 (partial graphics and performance upgrade)
- Enhanced existing carousel, striped coils, payout animation and factory campus; no duplicate facility.
- Closed DC and CLS roofs, translucent front facade, HVAC, rack components, DC component click descriptions, CTE/NMS and separately colored illustrative earth path.
- Shared primitive geometry, instanced repeated details and coils in WebGL, preserved CPU fallback. Pointer move coalesced into animation frame; paused actor updates skipped. Five quality modes and Auto hysteresis; rolling frame timing/renderer statistics API.
- Existing power equations, length, repeater logic and dry camera path retained. Existing manual CLS door visit remains available independently.
- Validation: JavaScript syntax compilation and exact-source checks on four protected functions only. Regression test updated/added but not executed; real browser screenshots, FPS, GPU/CPU/mobile measurements, full feature regression, LOD/texture compression, damping and full tour ship-stop sequencing remain open. No performance target claimed. New assets are procedural Three.js, not Blender exports.

## v3.0.1 — 2026-11-27
- Separate automatic ship/carousel surface view and underwater ROV view using continuous camera blends; no route or power equation change.
- Frame-based pointer damping; distance-based repeated-detail visibility, low-quality water geometry; CTE/NMS/power cabinet descriptions.
- Corrected stale v2.9 UI-label assertion (baseline 49/50, one pre-existing mismatch). Actual Three.js r128 plus original test harness executed in a JavaScript isolate with Node APIs and renderer substituted: 54/54 checks pass. Baseline unique scene geometry objects 1,732; v3.0 test 553 (not GPU measurements). Browser/GPU, screenshots, real mobile and 60/30 FPS targets remain unverified. Texture atlas/compression and full environment map/postprocessing quality tiers are not implemented.

## v3.1.0 — 2026-11-27
- Black jacket and regular yellow stripes on the entire wet route, retaining the two shore armor meshes. Illustrative color scheme, not a universal cable specification.
- Blue depth-dependent haze and lighting, visible overhead water, stronger shallow caustics, crossed light shafts, 20 stylized coral patches and six animated fish schools. Procedural merged meshes, not photographs or a surveyed marine habitat. Geometry count increases with detail; browser/GPU performance and appearance are unverified.
- Slow bounded vessel translation (±3 scene units, maximum 0.075 scene units/s at 1×), heave, tether interpolation and tracking automatic camera. Not a real ship speed or laying-rate model. Pausing freezes movement. Protected optical/power calculations and route preserved.
- 56/56 regression checks with actual Three.js r128 and substituted DOM/renderer/Node APIs; syntax checks passed. No real browser screenshot or FPS measurement.
