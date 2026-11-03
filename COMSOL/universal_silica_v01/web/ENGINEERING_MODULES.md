# Universal Silica Fiber Lab — Engineering Loss & PMD Modules

**Added 2026-11-03 · Phases 1→2→3 completed in order.**

These modules provide **parameter/measurement-anchored engineering estimation**, not full Maxwell or dopant-specific physics. Existing scalar LP01 FEM mode and G.654.E validation functionality remain separate and unchanged.

## 1. Attenuation / link budget

- **Engine:** [attenuation.mjs](./attenuation.mjs)
- **UI:** [attenuation_ui.mjs](./attenuation_ui.mjs)
- **Test:** [tests/test_attenuation.mjs](./tests/test_attenuation.mjs)
- **Inputs:** α1550, α1625 (dB/km), query λ ∈ [1550,1625] nm, length (km), Pin (dBm), connector and splice counts/losses.
- **Method:** linear interpolation of measured/specification anchor attenuation; no extrapolation beyond the anchors.
- **Outputs:** α(λ), fiber loss, joint losses, total link loss and Pout, curve of loss vs length, JSON download.
- **Validation example (illustrative, NOT a manufacturer measurement):** α1550=0.160 dB/km, α1625=0.175, λ=1550, length=80 km, 2 connectors ×0.2 dB, 1 splice ×0.05 dB, Pin=0 dBm ⇒ 12.8 dB fiber +0.4 connectors +0.05 splice ⇒ **13.25 dB total / −13.25 dBm output.**
- **Boundary:** no Rayleigh/OH physics identification, temperature effects, bending/PMD, or spurious extrapolation beyond measured wavelengths.

## 2. Macrobending loss (measured-anchor model)

- **Engine:** [bending_loss.mjs](./bending_loss.mjs)
- **UI:** [bending_ui.mjs](./bending_ui.mjs)
- **Test:** [tests/test_bending_loss.mjs](./tests/test_bending_loss.mjs)
- **Inputs:** two measured bending radii (mm) and four **positive measured loss-per-turn anchors** at 1550 and 1625 nm (2 radii × 2 wavelengths), query radius and wavelength, number of turns, straight-link loss.
- **Method:** log(loss/turn) interpolation in reciprocal radius 1/R, then log interpolation across λ. Strict **no extrapolation**. Additional bend loss is per-turn × turns, assuming additive turns.
- **Outputs:** dB/turn, total bend loss, bend length, straight-link + bend total, interpolated curve and JSON.
- **Validation example (illustrative, NOT a YOFC/LE125 measurement):** anchors at radii 10/20 mm; 1550 nm losses 0.15/0.005 dB/turn, 1625 nm losses 0.32/0.015 dB/turn; at 15 mm, 1550 nm, 10 turns ⇒ **0.015536 dB/turn**, **0.155362 dB additional bend loss**. If 13.25 dB baseline, combined **13.405362 dB**.
- **Boundary:** This is **not** scalar 2D/3D curved-waveguide FEM with PML or radiated-loss physics; do not call computed values product-level bending predictions. No zero-loss anchors in logarithmic interpolation; record the instrument detection limit and use justified positive floor if needed.

## 3. PMD / statistical DGD

- **Engine:** [pmd.mjs](./pmd.mjs)
- **UI:** [pmd_ui.mjs](./pmd_ui.mjs)
- **Test:** [tests/test_pmd.mjs](./tests/test_pmd.mjs)
- **Inputs:** PMD coefficient (ps/√km), length (km), coefficient **meaning**: mean-DGD/√km or RMS-DGD/√km.
- **Assumption:** strongly coupled, statistically stationary, isotropic random birefringent sections, resulting in an approximate **Maxwell DGD** distribution. DGD scales as √L.
- **Outputs:** mean, RMS, median, P95 and P99 DGD, Maxwell probability-density curve and JSON.
- **Validation example (illustrative):** mean DGD coefficient=0.05 ps/√km, 80 km ⇒ mean DGD=**0.447214 ps**, RMS=**0.485406 ps**, P95=**0.783433 ps**, P99=**0.943942 ps**.
- **Boundary:** A single-link DGD distribution quantile **is NOT** the manufacturer cable **PMD_Q** population specification. Real polarization-dependent n_eff and group delay, stress, twist, frequency-dependent Jones matrix and time evolution are not calculated in the present scalar LP01 FEM.

## 2+ Velamuri (2018) — bending physical-model bridge (2026-11-03)

Based on the official COMSOL 2018 paper [*Study of Bending Losses in Optical Fibers using COMSOL* (Velamuri et al.)](https://www.comsol.com/paper/study-of-bending-losses-in-optical-fibers-using-comsol-multiphysics-61111) / [published PDF](https://cn.comsol.com/paper/download/579211/velamuri_paper.pdf).

**New browser section:** `#bend-physics-tool`, between the prior macrobend measurement interpolation and PMD sections.

**New files:** [bend_physics_bridge.mjs](./bend_physics_bridge.mjs), [bend_physics_ui.mjs](./bend_physics_ui.mjs), [tests/test_bend_physics.mjs](./tests/test_bend_physics.mjs). These preserve the earlier measured-anchor `bending_loss.mjs` model.

### What is implemented now (transparent model levels)

1. **Equivalent straight-guide index visualization (conformal mapping):**
   `n_equiv(x) = n_stress(x)·(1+x/R)` for `|x| << R` (x is distance outward from bend neutral axis). `R` is **bend radius**, `d=2R` is **bend diameter**.
   - Unbent profile `n(x)` comes from current **Step / Graded / Trench / Multilayer** sidebar inputs and silica Sellmeier model.
   - Optional **illustrative elastic-beam proxy** (longitudinal strain `ε₁=x/R`, transverse strains `ε₂=ε₃=-ν ε₁`) applies the paper stress-optic law:
     `n_stress = n [1 - n²/2(P11 ε₁ + P12(ε₂+ε₃))]` with `P11=0.113`, `P12=0.252` reported in the paper.
   - **THIS IS NOT PAPER GEBT**: no structural mechanics solution or stress tensor from a real bent fiber. The graph visualizes an assumption and cannot directly produce a physical dB/turn prediction.

2. **Complex-mode eigenvalue postprocessing**: take **externally provided** `|Im(n_eff)|`, wavelength `λ`, bend diameter `d`, turns `N`:
   `loss_dB_per_m = (20/ln10)·(2π/λ)·|Im(n_eff)|` and
   `loss_dB_per_turn = loss_dB_per_m·π d`.
   Input wavelengths/diameters are converted to meters before arithmetic. This follows paper Eq. (3) and the `E(z)=E0 exp(iβz)` attenuation sign convention. The program accepts only the nonnegative **magnitude** to avoid different electromagnetic solvers' sign conventions. Verify mode sign convention upstream.

3. **External PML convergence diagnostics**: paste `PML_thickness_in_wavelengths,Im(n_eff)_magnitude` rows from two or more independent **complex eigenmode** runs. The tool converts values to dB/turn, compares the **last two** thicknesses to a selected relative change tolerance, and records a warning or pass. It does not run a PML calculation and the convergence criterion is a user-controlled diagnostic, not a standard.

4. **Display 7λ**: the paper found a `7λ` PML thickness effective for their study; at `1550 nm` the geometric thickness is `10.85 µm`. This is a cited paper example, **not** guaranteed to converge for every fiber. Actual verification must sweep PML thickness and mesh independently.

### Paper inverse example — not independent physics validation

Paper: bend-insensitive fiber (BIF), **10 mm bend diameter** (not radius), `1550 nm`: paper COMSOL simulated loss `0.0119 dB/turn` and experimental mean `0.012 dB/turn`. Inverting Eq. (3) gives an illustrative `|Im(n_eff)|≈1.07580760875×10⁻⁸`, which the new function converts back to `0.0119 dB/turn`. The **"논문 BIF 결과 역산 예시"** button demonstrates the formula only. It is **not** newly computed by the scalar FEM, not a reproduction of the paper's 2D simulation, and not a validation of YOFC or PremierLink LE125.

The paper modeled a `3.05 µm` core radius, `62.5 µm` cladding radius in its standard reference structure. Its bend-insensitive sample has an index trench and central dip; actual dopant and strain profiles are required before comparing to any specific manufactured fiber. The UI defaults to current sidebar geometry, not the original paper manufacturer's fabricated index map.

### Still required for independent curvature-loss prediction

- 2D transverse **complex eigenmode solver** with complex coordinate stretching / actual PML exterior and correct outgoing-wave boundary.
- Mesh + PML thickness + exterior domain convergence (radiation loss can be extremely sensitive to reflection and mode selection).
- **GEBT mechanics or calibrated stress/strain field** from structural solver, stress-optic anisotropy as appropriate, fiber coatings / finite cladding.
- Matched real index profile, actual bend **diameter** and test conditions, and comparison to measured dB/turn with uncertainty. The existing M2 vector mode code has **PEC boundary, no PML** and cannot independently provide leakage Im(n_eff).

For product comparison, explicitly label estimates **"Measured-anchor"**, **"Conformal-index illustration"** or **"External-complex-neff postprocessor"**. Do not label these **"2D COMSOL-equivalent FEM result"**.

## 2++ Numerical 2D scalar FEM + PML: real eigenmodes, NOT full-vector/GEBT

New stand-alone [`../bending_fem/solver.py`](../bending_fem/solver.py) **directly assembles P1 triangular FEM matrices** on a 2D cross section with a complex coordinate-stretched PML and solves the complex eigenvalue equation using SciPy. A [pinned Colab notebook](../bending_fem/Velamuri_Bend_2D_PML.ipynb) generates actual mode intensity maps, independently reruns mesh/PML sensitivity studies, and exports JSON for the newly added `#bend2d-fem` browser results display.

**Do not confuse three different bending functions:** (1) measured-anchor log interpolation (`bending_loss.mjs`), (2) conformal-index illustration and *external* complex index postprocessor (`bend_physics_bridge.mjs`), and (3) *direct* 2D complex eigensolver (`../bending_fem/solver.py`). The third is a substantial numerical capability upgrade but **remains scalar, uses a truncated numerical cladding window, and is neither GEBT nor the paper's full-vector COMSOL EWFD**.

The actual test example **fails** a user-selected 2% last-two-point relative change screening: **mesh 4.729% and PML 3.507%**. The quoted loss **67.556 dB/turn** is therefore **unconverged, not product-validated and not comparable to Velamuri's BIF 0.0119 dB/turn**.

**Full numerical report and caveats:** [../bending_fem/VALIDATION_REPORT.md](../bending_fem/VALIDATION_REPORT.md).

## Development/testing

Run in repository root (Node.js 18+):

```sh
node COMSOL/universal_silica_v01/web/tests/test_solver.mjs
node COMSOL/universal_silica_v01/web/tests/test_attenuation.mjs
node COMSOL/universal_silica_v01/web/tests/test_bending_loss.mjs
node COMSOL/universal_silica_v01/web/tests/test_pmd.mjs
node COMSOL/universal_silica_v01/web/tests/test_bend_physics.mjs
```

The initial UI numeric values in all 3 sections are **examples only**, not measured values of a specific commercial fiber. For a manufacturer comparison, first supply production/measurement data with measurement conventions and specimen conditions.

## Future fidelity upgrades (not yet implemented)

1. **Attenuation:** full spectral measured data and a calibrated Rayleigh + IR + OH model with identifiable coefficients, uncertainty bars and reference checks.
2. **Bending loss:** 2D full-vector or scalar curved-index eigenmode with PML, complex beta, curvature and PML convergence, then IEC 60793-1-47 mandrel-winding reference comparisons.
3. **PMD:** ellipticity/stress anisotropy vector eigenmodes and wavelength-dependent group birefringence, random Jones-section coupling, measured IEC 60793-1-48 / ITU-T G.650.2 reference comparisons.

## Test standards / context

- [ITU-T G.650.3](https://www.itu.int/epublications/publication/itu-t-g-650-3-2017-08-test-methods-for-installed-single-mode-optical-fibre-cable-links): installed link attenuation (includes splice/connector loss); PMD and spectral attenuation measurement context.
- [ITU-T G Suppl. 40 (2024)](https://www.itu.int/rec/dologin_pub.asp?id=T-REC-G.Sup40-202407-I%21%21PDF-E&lang=f&type=items): lists IEC 60793-1-47 macrobend-loss and IEC 60793-1-48 PMD measurement methods.
- [ITU-T G.657](https://www.itu.int/rec/dologin_pub.asp?id=T-REC-G.657-202408-I%21%21PDF-E&lang=e): practical macrobend measurement and number-of-turns conventions.
