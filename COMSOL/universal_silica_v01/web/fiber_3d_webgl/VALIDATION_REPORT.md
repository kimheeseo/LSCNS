# WebGL 3D Fiber Visualization — Validation record

Date: 2026-11-04 · Research prototype v0.2

## Test matrix

| Check | Status | Meaning |
|---|---|---|
| Node module syntax | PASS | `node --check` |
| 3D BPM sample construction | PASS | 9 BPM intensity slices -> finite XYZRGBA point cloud; no dummy light samples |
| 3D cylindrical solid geometry | PASS | 3D line segment vertices for core and cladding |
| 3D Nested Tube structure | PASS | Tube rings extruded at actual input XY and nested radii |
| Guard malformed BPM arrays | PASS | Rejects unexpected dimensions |
| Browser WebGL API wiring (mock) | PASS | Draw calls include 3D GL_LINES and GL_POINTS; no JS errors |
| Browser playback z slider | PASS | Reveal goes from 1 to 0.5 on middle slide |
| HCF switching | PASS | Geometry updated; 3D optical intensity point cloud cleared |
| Switch 3D / fallback 2D | PASS | Both modes selectable |
| Native GPU pixel rendering | NOT VERIFIED | Execution sandbox's Chromium blocks WebGL/WebGL2 contexts |
| Optical loss & full-vector 3D Maxwell | NOT IMPLEMENTED | These were never part of this visualization upgrade |

The data pipeline is **BPM 2+1D scalar `I(x,y,z)` → sampled 3D cloud → WebGL2**. No new optical eigenvalue, attenuation or field propagation results are fabricated. Browser animation speed and changing shader color phase are purely visual, unlike a measured optical carrier or a physically modeled pulse.

Existing `run3D` uses Split-step Fourier paraxial BPM with real refractive-index profile and periodic lateral boundaries. All previous 3D physics tests and prior Full-vector 2D eigenmode / HCF geometry modules remain separate.

## Reproduction

```bash
node COMSOL/universal_silica_v01/web/tests/test_fiber_3d_webgl.mjs
# User browser: open GitHub Pages /COMSOL/universal_silica_v01/web/#fiber-3d
```

For a genuine 3D Maxwell propagation model, independent Full-vector FDTD/FEM with appropriate absorbing boundaries, wavelength/material resolution and field continuity validation must be developed and benchmarked separately.
