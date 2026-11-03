# HCF Nested Tube Designer — Validation Report

Date: 2026-11-04 (KST). Module version: HCF analytical pre-designer v0.1.

## 1. Implementation verification

- New module integrated into existing Universal Silica Fiber web UI; previous SMF/MCF/PMD modules preserved.
- Automated Node.js built-in tests: **14/14 passed**. Tests include 4, 6, 8 and 12 outer capillaries with 1/2/3 ring levels; automated clearance, parent ring containment, first-order Sellmeier-corrected ARROW root residual, zero-crossing Bessel functions, independent finite-difference capillary dispersion check and output handling.
- Browser visual smoke test with headless Chromium on a locally injected document (the execution sandbox blocks local server navigation): **no JavaScript page errors**. Verified live edit 6x2 → 8x3; actual drawing updated from 12 to 24 glass annuli; auto-fit changed D₁ 24.96 → 19.580 µm; design A/B comparison populated and reported changes. Existing live hosted page deployment must be verified separately.

## 2. Default example numbers (not a product reference)

| Quantity | Computed | Source model |
|---|---:|---|
| Input core Ø | 32 µm | Design input |
| Number of outer tubes | 6 | Design input |
| Nested depth | 2 | Design input |
| D₁ outer tube Ø | 24.96 µm | Input D₁/Dc=0.78 |
| Adjacent outer-tube gap | 3.770 µm | Geometry |
| Outer tube angular coverage | 85.8214% | Geometric angle only |
| First tube 1st-order resonance | 1214.3145 nm | ARROW local wall equation, Sellmeier |
| First tube resonance frequency | 246.8821 THz | c/λ |
| Ideal capillary neff at 1550 nm | 0.999582563048 | J0 capillary formula |
| Ideal capillary Aeff | 383.275251 µm² | Bessel radial integral |
| Ideal capillary MFD, Petermann-II | 18.818344 µm | Bessel radial derivative integral |
| Ideal capillary D at 1550 nm | 2.9638548 ps/(nm·km) | Exact analytic second derivative, n_air constant |
| **Confinement loss** | **NOT COMPUTED** | Requires full-vector complex FEM + PML |

## 3. N sweep with automatic outer diameter fitting

| N | D₁ actual (µm) | Angular coverage | Minimum adjacent gap (µm) | Real confinement loss |
|---:|---:|---:|---:|---|
| 4 | 24.960 | 57.21% | 15.670 | Not computed |
| 6 | 24.960 | 85.82% | 3.770 | Not computed |
| 8 | 19.580 | 98.15% | 0.350 | Not computed |
| 10 | 14.028 | 97.49% | 0.350 | Not computed |
| 12 | 10.877 | 96.81% | 0.350 | Not computed |

The **increase in geometric coverage does not establish that confinement loss decreases** with N, and actual nested tube sizes vary when auto-fit is used. Comparisons at fixed D₁ with invalid intersecting tubes are rejected.

## 4. Scope and scientific validity

**Numerically valid:** geometric circle arrangement and rejection of overlaps, 1D-local ARROW wall resonance estimates, ideal constant-air-index capillary mode proxies, dimensional conversion from user-supplied Im(neff) to dB/km.

**Not validated and not claimed:** coupled antiresonant resonances and band widths for the actual 2/3 nested layers, HCF FEM field profile, tube-count-dependent confinement loss, real material/gas dispersion, bend loss, surface scattering or measured product loss. No 2D/full-vector FEM has been executed for this new module. The older separate bending_fem research sample does not have reliable mesh/PML convergence and is not treated as an HCF loss solver.

**Acceptance gate for future quantitative HCF losses:** matched silica/air geometry and wall thickness, full-vector complex neff, convergence vs mesh/PML/exterior domain, mode tracking, quantitative literature benchmark with explicit error bars, and experimental data for any claims about real manufactured fiber.
