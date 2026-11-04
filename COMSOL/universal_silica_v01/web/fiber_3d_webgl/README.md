# Universal Silica Fiber Lab — Live WebGL 3D Animation v0.2

[Open the existing 3D Fiber Lab](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#fiber-3d)

## What changed

Previously the viewer projected a tube and several BPM cross-sectional intensity snapshots onto a **2D Canvas**. It remains available under **2D 투영**.

The new **3D WebGL** mode uses a native, standalone WebGL2 renderer with:

- A genuine xyz camera with pointer-drag orbit, mouse-wheel zoom and optional auto-orbit.
- 3D cylindrical outlines for solid Step-index/Graded-index/Trench/Multilayer fibers. Glass boundaries are an illustrative extrusion; they are not meshed as 3D Maxwell finite elements.
- Actual **x,y,z intensity samples** from the previously computed **2+1D Scalar BPM** results. Consecutive BPM snapshots are linearly interpolated *for visualization*, displayed as an additive light cloud with a running phase/pulse animation.
- A z-playhead that progressively reveals the propagated light cloud as the original **▶ 애니메이션** button advances the BPM frame slider. The 2D transverse heatmap and MFD-vs-z chart stay visible.
- The previous **HCF nested-tube** designer's actual circles extruded as 3D tube walls. **HCF: geometry only.** It never renders simulated HCF light or invents confinement loss.
- A WebGL2 availability fallback to the original 2D Canvas visualization. No external Three.js CDN required; no GPU server or access token needed.
- A PNG button that exports whichever geometry viewport is active, plus existing numerical JSON/CSV exports.

### Use

1. Open the lab, scroll to **3D 광섬유 구조 및 광 전파**.
2. Select a solid silica profile and enter wavelength, radius, Δn, z length and mesh. Click **BPM 전파 계산**.
3. In the **3D WebGL** viewport drag to rotate, use the mouse wheel to zoom, and start **▶ 애니메이션** to watch the glowing propagation cloud grow along z. Use **↻ 자동 회전** to orbit the view.
4. Select **2D 투영** to compare against the older visualization. Choose **HCF/Nested Tube** to inspect nested glass tubes in 3D (no optical propagation/loss prediction).
5. Export a PNG / JSON / CSV. Browser-generated user results are saved **locally**, not automatically pushed to GitHub.

## Physical and numerical limitations

The light-cloud animation is a **3D visualization of 2+1D scalar paraxial BPM** computed in x/y at many z steps; it is *not a new 3D Maxwell eigenmode or 3D Full-vector FEM solution*. Interpolation between BPM snapshots is for display only; no extra fields are solved at interpolated z values. The optical carrier oscillation is far too rapid to display at actual optical frequency. The moving glow uses a visual pulse effect, **not a numerically modeled optical carrier or pulse propagation**.

The FFT transverse boundaries are **periodic**, not absorbing PML, and can cause wrap-around errors. The z-axis is deliberately compressed to display the long fiber on screen. Reflections, back-propagation, polarization, nonlinear effects, HCF/ARF/NANF confinement loss, real material attenuation, and all 3D Full-vector effects remain unsupported.

## Verification

- `node tests/test_fiber_3d_webgl.mjs`: checks volumetric interpolation, finite point intensities, two true xyz geometries, and invalid inputs.
- Browser DOM smoke test: WebGL API was mocked to test 3D mode control, draw calls, z reveal, HCF geometry and 2D fallback. This checks application wiring, **not actual GPU pixel output**.
- Headless Chromium in this execution sandbox has WebGL disabled by administrative policy, so real GPU/WebGL shader rendering could not be independently inspected. Production behavior depends on a browser that permits WebGL2.
- Core BPM physics and independent NumPy/JS tests are unchanged. See [`../propagation_3d/VALIDATION_REPORT.md`](../../propagation_3d/VALIDATION_REPORT.md).

## Source files

- [`../fiber_3d_webgl.mjs`](../fiber_3d_webgl.mjs) — native WebGL2 shader rendering, 3D extrusion geometry, BPM volume sampling, orbit/zoom, PNG and fallback.
- [`../fiber_3d_ui.mjs`](../fiber_3d_ui.mjs) — existing UI, connected via `mountFiberWebGL()` and the same BPM result data.
- [`../fiber_3d_core.mjs`](../fiber_3d_core.mjs) — existing physically computed scalar BPM engine, unchanged.

> Source is released as a **research visualization tool**, not a manufacturer-certified fiber propagation solver.
