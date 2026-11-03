# 2D complex-PML optical-fiber bending — research prototype

**Updated:** 2026-11-03  
**Reference:** Ashitosh Velamuri et al., [*Study of Bending Losses in Optical Fibers using COMSOL*](https://www.comsol.com/paper/study-of-bending-losses-in-optical-fibers-using-comsol-multiphysics-61111), COMSOL Conference 2018. [Official paper PDF](https://cn.comsol.com/paper/download/579211/velamuri_paper.pdf).

## What changed

Unlike the original web attenuation/bending-loss table interpolation and the intermediate `bend_physics_bridge.mjs` that accepts external complex effective indices, **[solver.py](./solver.py)** now computes complex eigenvalues **itself** using 2D scalar triangular P1 FEM with a complex-coordinate PML. It can produce a 2D eigenmode-intensity array, output `n_eff`, convert its imaginary part into bending loss, and execute new finite-element solves for mesh and PML thickness changes.

It is **not** Velamuri's **full-vector COMSOL EWFD + GEBT** calculation. This implementation does not reproduce the paper's measured BIF index profile, central dip, detailed polymer coating, anisotropic strain tensor or 125 µm cladding exterior. Therefore, **it does not validate YOFC or LS PremierLink LE125**, and all results must be labeled *experimental / not independently verified*.

### Paper equation → implementation

With bend **radius** R=d/2, measured from the neutral axis to x toward the outside of the bend:

```text
n_equiv(x,y) = n_stress(r) (1+x/R)
```

The paper's stress-optic relation is:

```text
n_stress = n [1 - 0.5 n² (P11 ε1 + P12 (ε2+ε3))]
P11=0.113, P12=0.252
```

**Current `stress_mode='beam_proxy'` is only an optional first-order isotropic beam/Poisson approximation**; the actual GEBT structural-mechanics calculation is missing. Default `stress_mode='none'` uses conformal geometry only.

Using complex stretches `s_x=1+i σ_x`, `s_y=1+i σ_y`, 2D triangular FEM discretizes the scalar generalized eigenproblem:

```text
∇⊥⋅[ diag(s_y/s_x, s_x/s_y) ∇⊥E ]
      + k0² n_equiv² s_x s_y E = β² s_x s_y E

β = k0 n_eff ; E(z) ∝ exp(iβz)
PML: σ_x = strength * (distance_into_PML / PML_thickness)^p  (and y)
```

A consistent triangular P1 FEM mass/stiffness matrix is assembled with centroid sampling of material/PML coefficients and outer Dirichlet boundary; the complex generalized problem is solved through **explicit generalized shift-invert**, `(A-σM)^-1 M` with SciPy ARPACK. This is essential: SciPy's standard `eigs(A,M=M,sigma)` has different positivity assumptions that do **not** apply to complex PML mass matrices.

If `Im(n_eff)≥0` for this field convention:

```text
loss_dB/m = (20/ln 10) (2π/λ_m) Im(n_eff)
loss_dB/turn = loss_dB/m × π × bend_diameter_m
```

Numerically **negative** imaginary effective index is *not* silently made positive; the bend-loss value becomes unavailable for such a mode.

## Use

- [**Google Colab: Velamuri_Bend_2D_PML.ipynb**](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/bending_fem/Velamuri_Bend_2D_PML.ipynb) — reproducible source pinned inside the notebook, 2D intensity maps for straight/bent cases, mesh & PML sensitivity, export JSON.
- [**Web workbench**](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/?v=20261103-complex-pml-1#bend2d-fem) — select **02++ 실제 2D FEM/PML 결과** and import a JSON exported by Colab.
- [**Executed demonstration JSON**](./example_validation.json) — includes **failure** of a user-selected 2% last-two-grid and PML sensitivity check.
- [**Python tests**](./tests/test_solver.py) — actual P1 matrix, 2D eigenmode, sign, straight case, and limited mesh/PML checks.
- [**Current analysis report**](./VALIDATION_REPORT.md).

To run locally:

```bash
pip install numpy scipy matplotlib pandas pytest
OPENBLAS_NUM_THREADS=1 python solver.py --core-radius-um 4.1 --delta-n-core 0.005 \
  --bend-diameter-mm 10 --window-um 22 --pml-thickness-um 10.85 --grid-h-um 0.7
OPENBLAS_NUM_THREADS=1 python solver.py --convergence
OPENBLAS_NUM_THREADS=1 python -m pytest tests/test_solver.py -q
```

The first example is a **made-up Step-index** design, **not the paper's bend-insensitive fiber**, and not manufacturer's spectral bending-loss specification.

## Executed demo and numerical reliability

Settings: λ=1550 nm, core radius=4.1 µm, Δn=0.005, bend **diameter=10 mm** (radius 5 mm), pure-silica computational window halfwidth 22 µm, 7λ=10.85 µm PML, mesh h≈0.7 µm, stress correction OFF.

| Output | Direct 2D P1 FEM/PML run |
|---|---:|
| Re(n_eff) | 1.446327526 |
| Im(n_eff) | 6.1073495×10⁻⁵ |
| Computed loss (**not converged**) | 67.55619 dB/turn |
| Straight-fiber `Im(n_eff)` | ≈−4.5×10⁻¹² (small numerical sign error; treated as near-zero only in straight check) |
| h=0.8 µm | 70.84181 dB/turn |
| h=0.65 µm | 72.18885 dB/turn |
| h=0.55 µm | 68.92921 dB/turn |
| Mesh last-two relative difference | **4.729% — FAIL vs 2% diagnostic** |
| PML 4λ=6.20 µm | 62.72457 dB/turn |
| PML 6λ=9.30 µm | 69.92529 dB/turn |
| PML 7λ=10.85 µm | 67.55619 dB/turn |
| PML last-two relative difference | **3.507% — FAIL vs 2% diagnostic** |

**Conclusion: The example has NOT met mesh/PML numerical convergence; its numeric loss is NOT a validated prediction.** 2% is a *user-chosen screening criterion*, **not a published acceptance standard**. Domain width and mode-continuation checks remain incomplete even if these two criteria pass.

These loss magnitudes differ dramatically from the paper's reported **0.0119 dB/turn** for the actual specially fabricated BIF at 1550 nm and 10 mm diameter, so they must **not** be used as evidence of paper reproduction or solver accuracy.

### Why 125 µm physical cladding matters

The PML starts at a truncated computational halfwindow (for the example, |x|,|y|=22 µm), **inside the physical 62.5 µm cladding radius of a standard optical fiber**. This is a numerical testbed; it does not model an actual physical glass/polymer interface or its radiation. The realistic boundary, coating, true RIP and mechanical field must be implemented and verified before the value can be linked to real fiber macrobending.

### Next research tasks (not yet completed)

1. Calibrate the paper's real **central-dip and trench refractive index profile** and actual bent-fiber coating/radial material stack.
2. Expand the physical domain through the actual 62.5 µm cladding radius, using adaptive mesh, and repeat mesh/PML strength/thickness/window convergence.
3. Extend scalar to **full-vector 2D electromagnetic FEM** with anisotropic permittivity, two polarization modes and correct outgoing radiation.
4. Replace stress proxy with **GEBT strain tensor** or validated mechanics solver, and compare against published BIF experiments and IEC 60793-1-47 tests.

**Important:** The new notebook numerically solves a 2D P1 finite-element eigenproblem; the GitHub Pages web app **only imports its output JSON**. It is misleading to claim that the browser's existing 1D LP₀₁ solver now predicts bend radiation by itself.
