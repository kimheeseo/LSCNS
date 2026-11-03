# Full-vector 2D FDFD & E-field Integral Validation — 2026-11-04

## Implemented vs proposed

- **DONE:** 2D coupled Hx/Hy vector eigenmodes, real neff, reconstructed complex Ex/Ey/Ez/Hz, full 2D E-field intensity map, MFD second-moment and near-field PII-like integral, electric-field Aeff integral, independent step-index exact HE11 neff oracle, mesh sensitivity tests, Colab and GitHub Pages results viewer.
- **NOT DONE:** Full-vector **FEM**, complex PML, air-core NANF/HCF loss, vector anisotropic PMF, material-weighted nonlinear Aeff, verified real manufacturer profile. The other scalar FEM solver remains available.

## Reference physics and validity

Maxwell H-field vector eigenproblem for translationally invariant isotropic waveguide. Boundary: Hx=Hy=0 on a finite Cartesian square. Interior index: n_silica(λ) + piecewise Δn. Eigenpairs from `scipy.sparse.linalg.eigs` on coupled 2N² blocks, finite differences. `neff=Re(sqrt(β²)/k0)` for guided solid-core mode. No complex PML: do not derive confinement-loss dB/km from tiny Im(neff) solver artifacts.

For step-index V<2.405, compare against the independent analytic **full-vector HE11** Bessel/modified-Bessel characteristic equation. Exact eigen equation reference: https://doi.org/10.1007/s10825-023-02006-y .

MFD conventions and test methods: IEC 60793-1-45, https://webstore.iec.ch/en/publication/3457 . This 2D intensity integral is a *numerical definition*, not a direct IEC measurement.

## Numerical outputs (identical illustrative design)

| Metric | Output |
|---|---:|
| λ | 1550 nm |
| a | 5.5 µm (assumed) |
| Δn | 0.003478579 (previous MFD fit) |
| finite computational halfwidth | 22 µm |
| FDFD grid | 69×69 |
| Exact HE11 neff | 1.445716906244325 |
| Vector FDFD neff | 1.445711888509771 |
| Absolute neff error | 5.01773455×10⁻⁶ |
| Vector MFD second moment | 12.791201 µm |
| Vector MFD PII-like | 12.55817 µm |
| Vector Aeff | 121.35505 µm² |
| longitudinal Ez fraction | 0.00074224 |
| finite-domain near-boundary intensity fraction | 2.3969×10⁻⁶ |

## Mesh study and caution

| Grid | Vector FDFD neff | Absolute error vs exact HE11 |
|---:|---:|---:|
| 35×35 | 1.44581069923 | 9.3793e-05 |
| 45×45 | 1.44575284630 | 3.5940e-05 |
| 57×57 | 1.44566072557 | 5.6181e-05 |
| 69×69 | 1.44571188851 | 5.0177e-06 |
| 81×81 | 1.44570171402 | 1.5192e-05 |

**Non-monotonic:** radial core boundary represented by staircase pixels; 69 grid happens to be closer to exact eigenvalue than 81. Numerical consistency in one mesh is **not convergence proof**. Further grid refinement, subpixel-permittivity treatment, conforming edge FEM, and independent domain checks remain mandatory for quantitative engineering certification.

## Executed tests

- `pytest tests/test_vector_modes.py`: **9 tests passed**. Covers Gaussian exact analytic `MFD=2w`, `Aeff=πw²`, scaling invariance, dual polarization HE11 degeneracy, exact HE11 neff sanity, guided metrics/field fractions, graded/trench/multilayer geometry, JSON modal-field schema and invalid inputs.
- Notebook: **8/8 executable cells completed** in Jupyter/Colab-compatible kernel; no Python exceptions. Static plots embedded in notebook output.
- Browser UI module: **Playwright + Chromium** using locally injected ES module and result JSON upload. Verified panel creation, JSON import, numerical cards `1.445711889 / 12.791 / 121.355`, Ex/Ey/Ez maps and no JS runtime exceptions. GitHub Pages live URL checked separately after push.

## Critical interpretation

- This code does **not** replace the existing scalar radial FEM, and the two methods have different discretization error. Scalar LP01 Bessel reference and vector HE11 reference are different electromagnetic models.
- For high-contrast or HCF structure, source code must be extended with complex radiation/PML boundaries and stronger discretization. Do not use the current `neff_imag_numerical` value to calculate dB/km.
- Manufacturer data matching cannot be claimed when delta n is *fitted* using published MFD. The current 2D vector Aeff is a **prediction for that assumed model**, not a validated LE125 specification.
- The web panel is a **snapshot importer**, not a server-hosted full-vector eigensolver. Run the Colab kernel after each design change and import its new JSON.
