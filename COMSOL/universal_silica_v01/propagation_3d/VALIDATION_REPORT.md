# 3D Fiber BPM — Validation Report (2026-11-04)

## Scope

The tool calculates **real numerical scalar 2+1D paraxial propagation**: a transverse `A(x,y)` field is propagated in z using split-step Fourier operators, producing a 3D sampled intensity volume. It does **not** solve Maxwell's equations in a 3D FEM/FDTD volume. The separate HCF 3D feature only extrudes the nested tube geometry.

## 1. Independent homogeneous Gaussian diffraction

Homogeneous silica, λ=1550 nm, initial scalar Gaussian field `exp(-r²/w0²)`, w0=6 µm, FFT 128×128, window ±45 µm, z=90 µm over 90 steps:

| Quantity | Result |
|---|---:|
| Exact paraxial Gaussian MFD `2w0 sqrt(1+(z/zR)²)` | 15.781789477930811 µm |
| Tested JS BPM MFD | 15.781789477930850 µm |
| Relative difference | approximately 2.7×10⁻¹⁵ |
| Max normalized power deviation | 2.63×10⁻¹³ |

The analytic comparison verifies the **paraxial diffraction discretization** for this test, not all real fiber physics.

## 2. G.654.E-like equivalent Step-index case

**Illustrative only**: assumed core radius 5.5 µm, Δn=0.003478579..., λ1550 nm, Gaussian launch waist4.5 µm, lateral offset1.5 µm, FFT grid128, halfwidth36 µm, 120 steps, z300 µm.

Local NumPy reference (independently executed in the implementation environment):

| Output | Value |
|---|---:|
| Final relative numerical power | 1.00000000000004 |
| Final 2nd-moment beam MFD | 23.68909454 µm |
| Final scalar beam Aeff | 169.57001663 µm² |
| Final edge energy (>85% window) | 0.293933% |
| Maximum numerical power error | 4.0×10⁻¹⁴ |

**Interpretation:** Final MFD and Aeff are of an evolving *offset Gaussian* at z=300 µm, **not** stationary LP01 modal MFD/Aeff nor independently validated manufacturer product specifications.

## 3. Numerical sensitivities (local NumPy reference)

| Grid | Longitudinal steps | Final MFD (µm) |
|---:|---:|---:|
| 64² | 120 | 23.378924 |
| 128² | 80 | 23.944113 |
| 128² | 120 | 23.689095 |
| 128² | 240 | 23.573657 |
| 256² | 120 | 23.796234 |

Mesh/step/domain convergence is **not certified**; numerical conservation does not imply physical accuracy. At long z, the FFT periodic transverse boundary may cause wrap-around. No PML is implemented.

## 4. Code checks

- Locally developed Python BPM engine: **4 pytest test groups passed** (homogeneous Gaussian, 4 solid profiles, guided-vs-free trend, invalid-input guards).
- Locally developed JavaScript BPM engine: **22 assertions passed**.
- Self-contained locally generated Colab notebook: **16 cells executed without errors**, 3D plot and output ZIP generated.
- Browser UI with corresponding local code was smoke tested in Chromium: actual BPM run, 3D view, slider, HCF geometry-only branch and export controls; no errors. Exact GitHub-integrated compact UI should be independently smoke tested after deployment.
- Compact GitHub browser core additionally directly executed in an isolated JS runtime: homogeneous Gaussian test passed and power deviation < 1e⁻⁹. Exact compact source may differ in code organization from fully tested local package; do not claim byte-identical implementation.

## 5. Not validated

**3D vector Maxwell FEM/FDTD, port reflections, nonparaxial fields, PMF polarization, propagation-based material attenuation, HCF/NANF full geometry-resolved confinement loss dB/km, stochastic MCF coupling and experimental products remain outside this implementation.** HCF tube count changes the geometry but not a predicted loss value.

Reference code: `../web/fiber_3d_core.mjs` and `propagation_3d.py`. The independent existing 2D Full-vector FDFD eigenmode solver is maintained separately.
