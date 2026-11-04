# HCF Complex Eigenmode FEM + Resonance Research v0.1

**2026-11-04 — the existing Universal Silica Fiber Lab now has a separate FEM results panel.**

- **[Existing live HCF designer & imported FEM results](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#hcf-fem)**
- **[Run the FEM solver in Google Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/hcf_complex_fem/HCF_Complex_FEM_Research.ipynb)**
- [Source code](hcf_complex_fem.py) · [Validation report](VALIDATION_REPORT.md) · [Example FEM spectra](results/default_spectrum.csv) · [Example FEM field](results/default_field.png)

## What has actually been implemented

1. **Nested HCF real 2D geometry:** 4–12 outer tubes, nesting depth 1–3, user-defined air core, tube diameter ratios, glass-wall thicknesses, and gaps. Python geometry definitions mirror the preexisting browser HCF designer (`web/hcf_designer.mjs`). Overlapping tube errors trigger input validation; auto-fit reduces D1 when configured.
2. **Numerically solved complex effective index:** two-dimensional **scalar** wave equation, P1 triangles, assembled sparse generalized eigenproblem with Cartesian complex-coordinate-stretched PML and shifted sparse eigenvalue solve. This is an actual FEM computation on the tube cross-section, NOT the old ideal capillary `neff` formula.
3. **Actual spectral re-solves:** for each wavelength the FEM assembles/re-solves the eigenproblem with wavelength-dependent pure silica Sellmeier dispersion; it saves `Re(neff)`, `Im(neff)` and PML-derived attenuation **preliminary values**. An ideal J0 mode-overlap filter flags unreliable mode selections.
4. **Candidate resonance/antiresonance detection:** candidates are **local maxima/minima of the FEM-derived preliminary attenuation** at the supplied discrete wavelengths. They are not fixed by ARROW and not a verified exact resonance. If the core mode cannot be identified at neighboring wavelengths, features are suppressed.
5. **Web UI integration:** the existing HCF Designer renders geometry and ARROW approximations live. The new `HCF FEM` panel exports that design's JSON, links to the executable Colab notebook, imports the FEM result, and shows complex neff, core fraction, preliminary loss trend, mode field and candidate extrema. **FEM is not calculated in JavaScript or continuously on parameter changes**. Input changes invalidate stale FEM result display.
6. **Research checks:** numerical unit regression tests, separate mesh/PML reruns, executed Colab notebook, and published default result JSON/CSV/plots.

## **Critical limitations — do not overinterpret**

This is **2D SCALAR FEM**, not full-vector edge-element Maxwell FEM. The P1 scalar wave equation cannot capture the proper HE/EH/TE/TM/vector boundary physics at air/silica interfaces of HCF. The Delaunay mesh samples tube-wall interfaces but is **not formally constrained/conforming**; thin-wall leakage calculations are highly mesh sensitive. It uses finite square PML + Dirichlet outer boundary and no air absorption or silica-surface scattering. No experimental or independent COMSOL HCF benchmark is yet matched.

**The default 1550 nm example fails the convergence study:** changing mesh spacing changes its preliminary leakage by ~499%, and changing PML thickness by ~69%. Therefore, **no result is approved as quantitatively reliable confinement loss in dB/km**. All filenames/fields intentionally use `UNVERIFIED` / `unverified`. The 22-point 750–1800 nm scan is far too coarse to identify narrow resonances. Some numerical candidates may be artifacts. `mode_identification_pass` describes J0 shape overlap only; it is not convergence validation.

## Calculation and equations

For a longitudinally uniform dielectric scalar field `psi(x,y,z)=psi(x,y) exp(i beta z)`, with convention exp(-i omega t), solve

\[
\nabla_\perp\cdot(\mathbf C\nabla_\perp\psi)+k_0^2\,n^2(x,y)\,s_xs_y\psi=\beta^2 s_xs_y\psi,
\qquad \mathbf C=\operatorname{diag}(s_y/s_x,\,s_x/s_y)
\]

where `s_x`/`s_y` are PML stretching factors (unity inside the physical domain); the P1 finite elements use consistent triangular mass and stiffness matrices. Then `n_eff=beta/k0`. With `Im(neff)>0` and the above propagation convention, preliminary confinement attenuation is:

\[
\alpha_{\mathrm{PML}}\,[\mathrm{dB/km}]
=\frac{40\pi\times 10^3}{\lambda_{\mathrm{m}}\ln 10}\operatorname{Im}(n_\mathrm{eff}).
\]

**This equation converts a PML-computed imaginary part; it does not guarantee that the imaginary part is physical or converged.** The ARROW thin-wall resonance equation remains available in the original HCF designer as an independent analytical guide.

## File index

| File | Description |
|---|---|
| [`hcf_complex_fem.py`](hcf_complex_fem.py) | HCF geometry, adaptive triangular mesh, complex P1 FEM/PML, mode selection, wavelength sweep and convergence study |
| [`HCF_Complex_FEM_Research.ipynb`](HCF_Complex_FEM_Research.ipynb) | Executed Colab notebook, upload designer JSON, actual FEM wavelength sweep, field map, CSV/JSON export |
| [`../web/hcf_complex_fem_ui.mjs`](../web/hcf_complex_fem_ui.mjs) | Integrated existing web designer export/import FEM result viewer |
| [`results/default_sweep.json`](results/default_sweep.json) | Real, **unconverged**, 22-point FEM run with saved scalar mode points at 1550 nm |
| [`results/default_convergence.json`](results/default_convergence.json) | Independent FEM reruns with mesh / PML perturbations, both failing the preliminary 5% criterion |
| [`results/default_spectrum.csv`](results/default_spectrum.csv) | Wavelength-by-wavelength output, all losses explicitly unverified |
| [`results/default_spectrum.png`](results/default_spectrum.png) | Illustrative complex FEM spectral plots |
| [`results/default_field.png`](results/default_field.png) | Real computed scalar field XY intensity, not a drawn Gaussian |
| [`tests/test_hcf_complex_fem.py`](tests/test_hcf_complex_fem.py) | Python tests for geometry, FEM complex system, spectral extrema and validation semantics |

## Reproduce locally

```bash
python -m pip install numpy scipy matplotlib pytest nbformat nbclient
OPENBLAS_NUM_THREADS=1 python -m pytest -q tests/test_hcf_complex_fem.py
OPENBLAS_NUM_THREADS=1 python hcf_complex_fem.py --field --output hcf_FEM_1550_example.json
OPENBLAS_NUM_THREADS=1 python hcf_complex_fem.py --sweep 800 900 1000 1100 1200 1300 1400 1500 1550 1600 --field --output hcf_FEM_sweep.json
```

Run Python commands from the directory containing this README, or execute the linked notebook in Colab. In the browser, new values require a new Colab run; current web UI is a result viewer, not a FEM backend.

## Follow-up for COMSOL equivalence

1. Implement **full-vector Maxwell Nédélec edge FEM** in 2D cross-section to correctly represent tangential field conditions and TE/TM/HE/EH modes.
2. Use genuinely **interface-conforming curved meshes** with controlled wall-element count, adaptive refinement and appropriate PML for open leaky modes.
3. Demonstrate independent mesh, domain, PML thickness, PML strength and mode-tracking convergence, ideally **less than 5% in dB/km** across multiple refinements, then verify with public COMSOL microstructured leaky-mode examples and published NANF data.
4. Only after passing validation, add automated nested-tube optimization and report production-quality 2D confinement loss predictions.
