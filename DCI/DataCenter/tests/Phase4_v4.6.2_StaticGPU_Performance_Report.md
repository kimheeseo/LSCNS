# LS Datacenter 3D v4.6.2 — Phase 4 GPU Static Geometry Cache & Rendering Optimization

Date: 2026-10-11 UTC. All measurements below derive from the same-run GitHub Actions workflow on Ubuntu headless Chromium + SwiftShader **software WebGL**. Not physical workstation or mobile GPU benchmarks.

## Scope and implementation

Existing v4.6.1 functionality is preserved; v4.6.2 adds:

1. **Static campus-ground geometry cache**: ground, road lines, perimeter and static landscaping geometry are recorded once and reused.
2. **Static facility-zone geometry cache**, keyed by zone, facility mode, roof visibility and section mode. Changes trigger cache lookup/rebuild rather than reusing stale geometry.
3. **Persistent GPU WebGL buffers** for static triangles, lines and transparent surfaces. An ordered signature of visible static geometry determines when the combined GPU vertex buffers must be re-uploaded. Camera motion within an unchanged visible-zone set does not re-upload these vertices. All dynamic elements (staff, animated systems, rack inspection, scenes, selected assets and route animations) remain on the existing batch buffers.
4. **Conservative route-segment frustum culling** for visible power/cooling/fiber topology paths. Supports the existing camera/worker POV; route segments are retained when within their bounding spheres.
5. **Live diagnostics**:
   - `window.LS3D_PHASE4_CACHE`: `enabled`, `entries`, `hits`, `misses`, `approxBytes`, `gpu.counts`, `gpu.uploads`, `gpu.bytes`, `clear()`.
   - `window.LS3D_RENDER_METRICS`: `triVertices` (dynamic), `staticVertices`, `wireVertices` (dynamic), `culledRouteSegments`, asset/rack culling stats, geometry/render CPU milliseconds.
6. No extra runtime libraries or replacements of the existing WebGL renderer. **GPU instancing is not implemented**: the renderer already batches geometry by opaque/wire/glass passes, and a risky shader rewrite was avoided pending a measured benefit.

The original `실측 CAD/BIM이 아닌 교육·설계 검토용 3D 개념 모델` and `가상값/가정값` disclaimers are retained.

## A/B benchmark (v4.6.1 vs v4.6.2 on same CI runner)

Source artifacts: [Phase 4 baseline/optimized screenshots and JSON](https://github.com/kimheeseo/LSCNS/actions/runs/38107752024/artifacts/11690615048).

| Scene | v4.6.1 | v4.6.2 | Change |
|---|---:|---:|---:|
| Campus / medium quality | 12.9 FPS | 13.6 FPS | +5.4% |
| Campus / low quality | 19.4 FPS | 19.3 FPS | −0.5% |

**Caution:** These are single short (~4.2s) samples on a shared CI runner, and differences of this size cannot establish a statistically significant FPS improvement. Earlier approaches that copied the static mesh back into JavaScript arrays caused a medium-quality regression; they were replaced with GPU-resident static buffers before finalizing. No claim of 30/60 FPS or physical-device improvement is made.

Measured geometry and cache at final medium-quality sample:
- v4.6.1 dynamic triangles: **87,348 vertices**; v4.6.2 dynamic triangles: **67,980 vertices** (**22.2% less** dynamic uploaded triangle data).
- v4.6.2 static triangles/lines: **19,672 vertices** retained in GPU buffers.
- Cached geometry: **7 entries**, **658 cache hits**, **7 misses**, around **1.57MB** of JavaScript array geometry on this CI sample.
- The campus overview frustum encompassed all 70 assets and all active route segments, so **0 assets and 0 route segments were culled** in this particular camera pose. Additional camera-angle benchmarks are required to quantify their benefit.

## Regression results

- [Phase 1 tests](https://github.com/kimheeseo/LSCNS/actions/runs/38107751991): **46/46 passed**, 0 runtime or console errors.
- [Phase 2 tests](https://github.com/kimheeseo/LSCNS/actions/runs/38107752081): **62/62 passed**, 0 runtime or console errors.
- Coverage includes: FOV units/ticks, zoom and mobile pinch, rack door/tray/cassette, procedural GPU LOD, power/cooling/fiber faults, CPO topology and optical reroute, MEP checklist persistence/CSV/events, worker/walking/POV, redundancy comparison, 2D and Word links.
- Tests use Chromium mobile emulation, **not** physical iPhone/Android testing.
- Screen images validate that the tested interactions render; **formal pixel-level image-difference testing was not performed**.

## Deployment details

Website: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?v=4.6.2

Source: `DCI/DataCenter/LS_Datacenter_3D.html`; existing Extensions JS/CSS and Phase2A/Phase2B JS/CSS are unchanged but all `?v=` query strings have been bumped to the v4.6.2 release identifier.

## Remaining engineering tasks

1. **Real-device GPU benchmark**: Windows GPU, browser hardware acceleration, Android and iOS Safari. These are untested.
2. **Visual parity tests** across orbit angles/roof/section/MPO rack close-ups to ensure static GPU separation remains visually faithful (current interaction regression suite passed).
3. **Targeted frustum-culling test** at close or reverse-angle cameras, since the sample overview had zero culled assets.
4. **Optional true GPU instancing** for repeated rack/server silhouettes, only if GPU profiling justifies additional shaders/attribute buffers.
5. Measured engineering-level CPO port/BOM, MEP clearance and collision checks are out of scope for this performance phase.
