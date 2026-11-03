# HCF / Nested Tube Designer v0.1 — Universal Silica Fiber Lab

**Integrated Live UI:** https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#hcf-designer

A dynamic **HCF/NANF cross-section and analytical optical pre-design module** embedded within the existing Universal Silica Fiber Lab. It preserves the existing solid-core radial FEM, 2-core MCF coupling demonstration, and PMD/DGD statistics.

## Features

| Input | What visibly/recomputationally changes |
|---|---|
| Air core diameter 16–120 µm | Actual drawn core radius; ideal capillary neff, MFD, Aeff, waveguide dispersion |
| 4–12 cladding tubes | Number and angular position of all outer and nested tube rings; angular coverage and minimum gap |
| Nesting depth 1/2/3 | Inner circles offset toward the outside of the parent ring; geometry and per-wall resonance table |
| Outer tube/core diameter ratio | Tube sizes and clearances; optional collision-avoiding auto-fit |
| Nested D₂/D₁, D₃/D₂ | Nested tube sizes and positions, with physical-containment checks |
| t₁/t₂/t₃ silica wall thickness | Silica wall geometry and wall ARROW resonance / antiresonance wavelength estimates |
| λ, n_air | Ideal capillary proxy, wall phase/detuning and spectral dispersion |
| External independently computed complex neff | Conversion of user-provided |Im(neff)| to dB/km (NOT prediction) |

The UI additionally shows **N=4–12 geometric coverage sweeps**, a wall-by-wall resonance-detuning spectrum, ideal capillary `D(λ)`, A/B design comparison, and JSON / CSV / cross-section SVG exports. Changed designs are recomputed **in-browser**, without a Python server.

## Physics: clear separation of what is calculated

**1. Actual 2D geometry (not FEM).** `R_outer_center = R_core + R_outer_tube + gap_core`; neighboring tube gap is `2 R_outer_center sin(π/N) − D_outer`. With `autoFit` enabled, the *actual tube diameter is reported* after reducing an oversized requested value to avoid overlaps. Inner tube centers are displaced radially **outward** within their parent capillary with a user-defined air clearance. These are conceptual geometries and not fabrication-ready capillary attachment/contact geometries.

**2. Local thin-wall ARROW formula, not a global coupled-mode solution.** For each silica wall `t_j`, solve iteratively for `λ` using wavelength-dependent pure-silica Sellmeier:

`λ_res,m ≈ 2 t_j sqrt(n_silica(λ)^2 − n_air^2)/m`, `m = 1,2,3`.

The analogous antireflection / antiresonance center is `λ_AR,m ≈ 2 t_j sqrt(n_silica(λ)^2 − n_air^2)/(m − 1/2)`. A plotted detuning `abs(sin(π·2t_j sqrt(n_silica²−n_air²)/λ))` is a **dimensionless wall-phase diagnostic, NOT transmission, reflectivity, or loss**. Real nested-tube resonances include inter-tube coupling and are not generally equal to these independent-wall estimates.

**3. Ideal J0 hollow capillary proxy, not HCF vector FEM.** Use `u01=2.4048255577`, core radius `R`, constant air index `n_air`, and `neff=sqrt(n_air²−[u01 λ/(2πR)]²)`. Report ideal-capillary Petermann-II and second-moment MFD, intensity-derived Aeff (radial numerical integration) and waveguide D from the exact second derivative. **Tube count, tube ratios, wall thicknesses and nesting do not affect these proxy values**. Any claim that this module proves N-dependent optical confinement would be incorrect.

**4. No directly predicted confinement loss.** No 2D full-vector complex eigenmodes, outgoing PML, surface scattering, bend loss or measured dB/km exist in this module. The GUI displays **미계산**. If the user supplies *converged external* `|Im(neff)|`, the separate unit-conversion calculator reports `α[dB/km] = 40π×1000·|Im(neff)|/[λ(m) ln(10)]` (assuming `E~exp(iβz)` attenuation convention). The result is explicitly user-provided FEM **postprocessing**, not a result of this geometry model.

## Example: default design

- Six outer capillaries; two nested rings; air core diameter 32 µm; D₁/Dc=0.78; D₂/D₁=0.61.
- Walls 0.58/0.45 µm; minimum adjacent gap ≈3.77 µm; angular coverage ≈85.82%.
- 1550 nm capillary-only baseline: neff ≈0.999582563, Aeff≈383.275 µm², Petermann-II MFD≈18.818 µm, ideal waveguide D≈+2.964 ps/(nm·km).
- First-wall first-order resonance ≈1214.3 nm (≈246.9 THz), with silica dispersion accounted for.
- The above are *assumed geometry and ideal reference* numbers, **not tested HCF product measurements or nested-tube confinement-loss predictions**.

## Tests and future steps

Run `node COMSOL/universal_silica_v01/web/tests/test_hcf_designer.mjs` from the repository root. Node built-in `node:test` verifies geometry, no collision, nested containment, Sellmeier, dispersive-wall resonance, ideal capillary mode checks, spectrum and lack of fictitious loss. See [VALIDATION_REPORT.md](VALIDATION_REPORT.md).

To obtain defensible N/diameter-dependent dB/km:

1. Generate boolean-robust conforming silica/air 2D cross-section meshes with sub-wall resolution.
2. Add full-vector (edge-element) Maxwell complex eigenmodes and outward PML, identifying the air-core fundamental mode by core power overlap and mode continuity.
3. Converge mesh, domain, PML thickness, absorption profile, and find similar values with an **independent reference solver or published geometry**.
4. Include glass overlap, surface scattering, gas absorption, fabrication imperfections, and bending only after adding their validated physics.

## References

- Poletti, **Nested antiresonant nodeless hollow core fiber**, *Optics Express* 22 (2014): https://pubmed.ncbi.nlm.nih.gov/25321960/
- **Design of 2 μm Low-Loss Hollow-Core Anti-Resonant Fibers**, *Micromachines* 14 (2023): https://pmc.ncbi.nlm.nih.gov/articles/PMC10305596/
- **Hollow-Core Optical Fibers for Telecommunications and Data Transmission**, *Applied Sciences* 13 (2023): https://www.mdpi.com/2076-3417/13/19/10699
- Existing independent M8 capillary approximate reference: `COMSOL/COMSOL_Install/waveoptics_m8/src/mcf_hcf.py`.
