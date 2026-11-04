## v0.3 visible playback upgrade — 2026-11-04

- 연속적인 z 재생 위치를 셰이더에서 사용하여 발광 강도 선단 및 잔광을 표시. 입력 광장은 기존 BPM 계산된 snapshots만 사용.
- 단면 링은 실제 입구 XY 형상으로부터 추출하여 z 방향으로 이동. 위치 표시이지 새로운 물리장 또는 경계조건이 아님.
- 기존 700 ms setInterval 슬라이드 대신 requestAnimationFrame, 3/6/12초 재생과 계산 완료 후 자동 시작.
- 웹 단위 검사에서 광장 형상, moving-ring geometry, HCF geometry 및 UI 연결 확인.
- 실제 소비자 GPU의 픽셀 품질/프레임 속도는 별도 기기 검증이 필요. 기존 Web Smoke CI 실패는 별도의 COMSOL_Install/web/smoke.cjs 단계에서 발생했으며 이 3D 검증과 구분해야 함.

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
