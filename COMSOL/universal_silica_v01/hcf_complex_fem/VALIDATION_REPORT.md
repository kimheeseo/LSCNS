# Validation Report — HCF Complex Scalar P1 FEM + PML v0.1

**Report date:** 2026-11-04. **This is a prototype FEM upgrade; there are no validated quantitative HCF loss results yet.**

## Implemented

- Actual two-dimensional scalar P1 triangular finite-element eigenproblem with a complex-coordinate stretched Cartesian PML.
- Nested HCF design from the same geometric input keys as the web HCF Designer, 4–12 tubes and 1–3 nesting layers.
- Mode selection by core-localized ideal J0 scalar overlap and physical core-power fraction; weak overlap is flagged and excluded from FEM spectral extrema.
- Wavelength-by-wavelength reassembly/re-solve, complex `neff`, scalar field intensity and local loss-extrema candidates.
- Executed Google Colab notebook, 10 Python tests passing; browser FEM results import panel.

## Default research case

| Parameter | Value |
|---|---:|
| Air core diameter | 32 μm |
| Outer tubes × nesting | 6 × 2 |
| D1/Dcore, D2/D1 | 0.78, 0.61 |
| Wall t1/t2 | 0.58/0.45 μm |
| λ | 1550 nm |
| n_air | 1.00027 |
| FEM triangular nodes/elements | 6874 / 13502 |
| `Re(neff)` (scalar FEM) | **0.9997047421** |
| `Im(neff)` (scalar FEM/PML) | **3.4405 × 10^-12** |
| Air core scalar power fraction | **99.265%** |
| Scalar J0 overlap (mode identification only) | **0.9739** |
| Derived attenuation (dB/km) | **0.1211 — UNVERIFIED / NOT CONVERGED** |

The result is **not** a measured result, not an exact HE11 vector mode, not a vendor HCF data sheet and not a COMSOL benchmark reproduction.

## Independent convergence reruns (FAIL)

The solver was rerun, rather than interpolating earlier results, with 1) a finer background/wall mesh and 2) a thicker PML.

| Check | Percentage difference in preliminary loss | Criterion | Result |
|---|---:|---:|---|
| Mesh change | **~499%** | <5% | **FAIL** |
| PML thickness change | **~69%** | <5% | **FAIL** |
| Domain size, PML exponent/strength sweep | Not exhaustive | Required | Not verified |
| High-order mode tracking across resonance | Only J0 overlap | Required | Incomplete |
| Published COMSOL complex HCF neff / loss benchmark | Not reproduced | Required | Not verified |

Even if mesh/PML variations were below 5%, two-point convergence would be an initial screen only. Domain-size convergence, mode identity and agreement with published data are independently required.

## Resonance and antiresonance

The demonstration actually reruns FEM at 22 wavelengths from 750 to 1800 nm in 50 nm steps. Local maxima/minima of preliminary dB/km are reported as **candidates only**. At 1100 nm, the J0 overlap falls far below the minimum required for tracking, so peak detection suppresses candidates relying on that point. Several small ripples elsewhere may be PML/mesh artifacts, **not physical antiresonances**. Current sweep step is too coarse for quantitative resonant wavelengths.

The original web `ARROW` thin-wall wall resonance estimate still exists and is appropriately labeled a separate analytical model. FEM extrema are not substituted with that analytical estimate.

## Why this cannot yet claim COMSOL-quality HCF performance

The prototype uses **scalar** FEM and an approximate Delaunay wall mesh, while correct HCF requires full-vector Maxwell modes and carefully interface-conforming curved wall meshes. The imaginary part of leaky-mode neff is extremely sensitive to PML settings and discretization error. This demo establishes **execution of the complex FEM eigenproblem and a working data workflow**, not accurate leakage values.

## Automation and reproduction

```bash
OPENBLAS_NUM_THREADS=1 python -m pytest -q tests/test_hcf_complex_fem.py
node --check ../web/hcf_complex_fem_ui.mjs
```

- 10 numerical Python unit/regression tests passed in development environment.
- The 13-cell Colab notebook was executed end to end with `nbclient` without errors.
- All result JSONs explicitly set `converged_and_validated=false`; UI labels dB/km as unverified. Results are invalidated when current browser HCF dimensions change.
- Consumer browser import & plot need final deployment smoke; check GitHub Pages and CI status separately.

See [source](hcf_complex_fem.py), [default computed results](results/default_sweep.json) and [Colab](HCF_Complex_FEM_Research.ipynb).
