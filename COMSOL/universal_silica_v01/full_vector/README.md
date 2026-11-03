# Full-vector Fiber Mode Lab v0.1 — 2D FDFD + E-field integration

[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/full_vector/LS_Full_Vector_Fiber_v0_1.ipynb)

**Workflow:** Colab real vector eigensolver → compute Hx, Hy and reconstruct H_z, Ex, Ey, Ez → numerically integrate E-field intensity for **MFD, Aeff** → download JSON + NPZ → visualize the saved result in [the existing Universal Silica Fiber Lab](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#full-vector-tool).

This is an explicit **full-vector 2D finite-difference (FDFD) eigensolver**, not a scalar solver and **not full-vector FEM**. It is implemented as a separate Python numerical module, so the original JS radial scalar FEM and MCF/HCF/PMD modules continue unchanged. GitHub Pages renders imported results; it does NOT solve full-vector eigenproblems directly in the browser.

## Components

| Path | Role |
|---|---|
| `vector_modes.py` | Coupled 2D magnetic vector-field eigenproblem, real-index silica profiles, reconstructed electric field, 2D numerical integrals, exact step-index vector HE11 reference |
| `LS_Full_Vector_Fiber_v0_1.ipynb` | Offline-self-contained **executed** Google Colab notebook, user inputs, vector field visualizations, mesh check and ZIP/JSON exports |
| `../web/vector_mode_ui.mjs` | Integrated results viewer in existing GitHub Pages tool |
| `results/vector_mode_result.json` | Example full-vector result and downsampled Ex/Ey/Ez 2D intensity maps, selectable in web tool |
| `results/vector_fields_complex.npz` | Full complex vector components for first mode (Ex,Ey,Ez,Hx,Hy,Hz), index map and axis |
| `results/vector_metrics.csv` | Integrated scalar metrics from vector electric field |
| `results/vector_mesh_convergence.csv` | 45/57/69/81 Cartesian-grid case and exact HE11 neff comparison |
| `results/vector_Ex_Ey_Ez.png` | Actual vector-field component intensities |
| `results/vector_index_and_total.png` | Index section and modal intensity |
| `tests/test_vector_modes.py` | Numerical regression checks |

## Equations and material model

Waveguide is z-invariant, isotropic, nonmagnetic, lossless with `exp(iβz − iωt)` convention. Solve the **coupled transverse H vector** eigenproblem (derived from `curl(ε⁻¹ curl H)=k₀² H` and `div H=0`) for the eigenvalue `β²` on a 2D Cartesian FDFD grid. Reconstruct `H_z` using solenoidal constraint and `Ex/Ey/Ez` from Maxwell curl H. A small imaginary numerical eigensolution part is not a physically computed leakage coefficient: there is **no PML**.

Define `I = |Ex|²+|Ey|²+|Ez|²`. With centroids `xc,yc`:

- `Aeff = (∫ I dxdy)² / ∫ I² dxdy`.
- `MFD_2nd = 2 sqrt[2 (∫((x−xc)²+(y−yc)²) I dxdy)/(∫ I dxdy)]`.
- `MFD_PII-like = 2 sqrt[2 ∫I dxdy / ∫ |grad sqrt(I)|² dxdy]`. **This generalized 2D near-field estimate is not automatically equivalent to the far-field Petermann-II standard.**

All area integrals use micrometre Cartesian cells; Aeff is in µm² and MFD in µm. Silica Sellmeier plus wavelength-independent profile `Δn` approximates doped regions. Aeff is electric-intensity based; a full material-weighted nonlinear effective area is a different quantity and is not computed here.

## Default validation case

Step-index **illustrative G.654.E-like**, λ=1550 nm, assumed a=5.5 µm, Δn=0.0034785793846155, domain halfwidth 22 µm, 69×69 grid.

| Metric | Result |
|---|---:|
| Exact independent **full-vector HE11** effective index | 1.44571690624 |
| 2D **full-vector FDFD** effective index | 1.44571188851 |
| Absolute neff difference | 5.018×10⁻⁶ |
| MFD second moment (2D E-field) | 12.79120 µm |
| MFD 2D near-field PII-like | 12.55817 µm |
| Aeff (2D E-field) | 121.35505 µm² |
| Normalized Ez intensity fraction | 0.0742% |

The solver's **grid convergence is not monotonic** due to stair-step interface approximation. See [validation](VALIDATION_REPORT.md). The nominal 12.5 µm Petermann II MFD used to fit Δn belongs to an *illustrative scalar step-index model*; it is **not** a measured manufacturer profile.

## Run locally

```sh
python -m pip install numpy scipy matplotlib pandas pytest nbformat nbclient
OPENBLAS_NUM_THREADS=1 python vector_modes.py 69
OPENBLAS_NUM_THREADS=1 python -m pytest -q tests/test_vector_modes.py
```

Use the Colab notebook for visual outputs. `vector_modes.py` is the authoritative implementation; the executed notebook embeds a pinned copy of the exact source to allow offline reproduction.

## Accuracy boundary and upgrade path

1. Current **full-vector FDFD** uses Cartesian grid, with Dirichlet outer boundary; no PML, leaky modes, or birefringence material tensor.
2. Only solid, circular/concentric silica fibers: Step-index, Graded-index, Trench-assisted and Multilayer. Do not use HCF/NANF air-core cases; HCF modes are leaky and need complex/PML boundary plus careful mode identification.
3. MFD is a **defined vector-intensity integral**, not a universal manufacturer measurement. Verify instrument conventions per IEC 60793-1-45.
4. Next for FEM equivalence: Nédélec (edge) FEM with conforming core boundaries; HCF full-vector complex PML, repeated mesh/domain/PML convergence, measured reference.
5. For polarized/PMF products: extend permittivity to anisotropic tensor, resolve full Jones/propagation, and validate birefringence against measurements.

**This is a functional vector-mode upgrade and a credible numerical prototype, not COMSOL-equivalent validated FEM.**
