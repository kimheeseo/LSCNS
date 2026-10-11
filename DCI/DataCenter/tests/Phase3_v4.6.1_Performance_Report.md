# LS Datacenter 3D v4.6.1 · Phase 3 rendering optimization

## Result (2026-10-11)

Optimized version: **v4.6.1**, commit [c995123cd261dc8a4d29f32af427f6e5c3b85c83](https://github.com/kimheeseo/LSCNS/commit/c995123cd261dc8a4d29f32af427f6e5c3b85c83).

Baseline: **v4.6.0**, commit [9dce7156ab9f6642f948cd814e85af1f61975f8f](https://github.com/kimheeseo/LSCNS/commit/9dce7156ab9f6642f948cd814e85af1f61975f8f).

[Same-run benchmark and screenshots](https://github.com/kimheeseo/LSCNS/actions/runs/38107036284/artifacts/11689393655) · [Benchmark workflow](https://github.com/kimheeseo/LSCNS/actions/runs/38107036284).

### Same-run A/B (Ubuntu GitHub Actions, headless Chromium, SwiftShader software WebGL)

| Quality | Baseline v4.6.0 | Optimized v4.6.1 | FPS change | Baseline median frame time | Optimized median frame time |
|---|---:|---:|---:|---:|---:|
| Medium | 12.7 FPS | 17.2 FPS | +35.4% | 83.3 ms | 50.1 ms |
| Low | 13.4 FPS | 25.1 FPS | +87.3% | 66.7 ms | 33.4 ms |

The measurement compares both versions sequentially on the **same CI runner**, rather than claiming comparable FPS from different days/runs. Samples are approximately 4.2 seconds each following a brief warm-up. This is a single comparative run under software WebGL, **not the FPS of any real workstation GPU or phone**, and not a guaranteed improvement on all hardware. The low-quality screenshot and medium-quality screenshot are available with the benchmark artifact.

### Geometry observations

- In the sampled campus camera view: **31 simplified rack LOD instances**, **1 full-detail rack**, **5 simplified distant staff**, and **87,348 triangle vertices** for v4.6.1.
- **0 asset instances were frustum-culled in the initial wide campus camera view** because the entire area fell inside the conservative camera frustum. Frustum culling is installed for narrower field-of-view and lower-angle exploration, but its additional performance benefit needs targeted camera-orientation benchmarking.
- The existing renderer already used large vertex arrays / batched draw calls for solid triangles, lines, and transparent triangles. v4.6.1 **does not add hardware instancing**; it reduces CPU mesh creation costs and reuses Float32Array upload allocations.

### Changes

1. **Conservative frustum visibility**: broad-phase sphere test for asset/zone/NPC rendering. Does not remove any assets from selection, facility calculations, collision checks or event logs.
2. **Screen-space rack LOD**: small/distant racks draw simplified box silhouettes and the top ToR distinction, medium-size racks retain server fronts, and near racks preserve original detailed draw logic; open rack mechanical inspector is never simplified.
3. **NPC LOD**: distant staff render as simplified silhouettes, but full personnel models remain near the camera.
4. **Material/color cache**: hex-color parsing output reused for static opaque material colors.
5. **Reusable typed WebGL staging buffers**: triangle, line and transparent geometry share the existing three-call pattern while avoiding repeated upload-array allocations.
6. **Render diagnostics**: `window.LS3D_RENDER_METRICS` reports visible/culled items, LOD counts, vertex counts, and CPU render/geometry time.
7. **Cache bust/version**: HTML title/header v4.6.1 and `?v=20261011-461-perf` for unchanged extensions/Phase 2 assets.

### Regression verification

- [Phase 1 browser regression](https://github.com/kimheeseo/LSCNS/actions/runs/38107036292): **46 PASS, 0 FAIL, 0 JS errors**.
- [Phase 2 browser regression](https://github.com/kimheeseo/LSCNS/actions/runs/38107036274): **62 PASS, 0 FAIL, 0 JS errors**.
- These automate: FOV ruler, wheel/pinch, rack door/server tray/cassette, GPU semantic LOD, power/cooling/network/fiber scenarios, N/N+1/2N redundancy, worker POV, walk mode, CSV, MEP progress/event persistence, optical CPO/pluggable, fault/reroute, and mobile Chromium viewport geometry.
- The Phase 1+2 test suite verified compatibility, **not pixel-perfect parity** across the entire scene.

### Remaining work

- Target hardware benchmarking: Chrome/Edge on Windows with discrete/integrated GPUs; actual Android and iPhone Safari.
- Additional campus performance: zone-level/static-object mesh caching, careful draw-call and GPU timing investigation, route-segment culling, and a higher fidelity visual regression harness. Do not assume that 30/60 FPS is met on software WebGL.
- More accurate networking BOM/CPO port and thermal modelling depends on official hardware specifications.
- Continue displaying **교육·설계 검토용 개념 모델, 가상값/가정값, 실측 CAD/BIM 아님** in all future variants.
