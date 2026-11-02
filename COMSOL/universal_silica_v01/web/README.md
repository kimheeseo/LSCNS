# Universal Silica Fiber Workbench — Web UI v0.2

**Live browser app:** https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/

A standalone engineering-style interface for the existing `universal_radial.py` radial scalar FEM model. GitHub Pages serves static HTML/CSS/JavaScript and **runs actual finite-element eigenmode calculations in the browser**. It does not simply interpolate precomputed output. No external Python server, COMSOL software or GitHub credentials are needed.

## What the UI does

- Choose **Step-index**, **Graded-index**, **Trench-assisted**, or **Multilayer** circular silica profiles. Edit radius, wavelength, core index step, graded exponent, trench dimensions, concentric layer definitions, mesh resolution and nonlinear index.
- Click **모드 계산 실행** to solve the generalized finite-element eigenproblem for the radial weak-guidance LP0 fundamental mode. Show `neff`, Petermann-II and second-moment MFD, `Aeff`, `γ` and five-point **FEM-derived dispersion D**.
- View a 2D *rotational visualization of the radial mode*, the refractive-index profile, radial intensity, and a recomputed spectral dispersion plot. This is **not** a separately solved 2D Maxwell field.
- Sweep wavelength and independently recompute the FEM at every point. Compare results at multiple mesh resolutions. Export calculation JSON, CSV, PNG, a reusable input-design JSON. Import design JSON.
- **Input change invalidation:** previous result cards are cleared rather than presented as results for the changed input. Users must run the solver again.

## How to use

1. Open the [Web Workbench](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/).
2. Select a preset or set profile parameters in the left panel.
3. Click **모드 계산 실행**. Review the four headline quantities and the three primary figures.
4. Click **파장 스윕** to regenerate `D(λ)`, or **메시 수렴** for convergence checks.
5. Use **JSON 결과**, **CSV 결과**, **모드 PNG**, **설계 JSON** to download computed data. **These export files are downloaded locally, not automatically committed to GitHub.** Avoid placing a GitHub token in the webpage. To version personalized results in the repository, commit the downloaded JSON/CSV through GitHub or ask ChatGPT with the saved files.

## Exact model and limits

- **Solver:** axisymmetric scalar P1 finite elements of `∇²⊥ψ + k0² n² ψ = β² ψ`, transformed to radial form and solved as generalized tridiagonal eigenproblem via shifted inverse iteration. Natural regularity at `r=0`, Dirichlet field at outer computational radius.
- **Materials:** pure silica Malitson Sellmeier `n_cladding(λ)`, other layers are **wavelength-independent Δn offsets**. This is **not** a calibrated dopant-dependent Ge/F dispersion model.
- **Outputs:** eigenmode metrics are scalar approximations. Dispersion is a five-point finite difference of newly solved eigenvalues on identical radial geometry, not a fitted value or a manufacturer's measured dispersion. `γ` uses input `n₂` and computed `Aeff`.
- **Not supported:** arbitrary 2D cross sections, MCF, microstructured PCF, ARF/NANF HCF, full-vector electromagnetic modes, polarization, birefringence, true attenuation, bending loss, confinement loss, PML, and model inversion of unknown manufacturer profiles.
- **Domain:** model calculator radius is a numerical boundary, **not** actual 125 µm fiber cladding. Edge-truncation and mesh convergence should be checked for research use. Small solver residual does not establish physical accuracy.
- A G.654.E-like preset uses **assumed 5.5 µm core radius and fitted Δn** to target MFD ≈12.5 µm. This is not independent LE125 certification or COMSOL license-equivalent validation.

## Files and verification

- `index.html` (self-contained CSS), `app.mjs`: user interface.
- `solver.mjs`: independent JS implementation of Python radial FEM assembly/integration/eigenmode calculation.
- [`results/baseline_summary.csv`](results/baseline_summary.csv): five preset default calculations.
- [`results/baseline_verification.json`](results/baseline_verification.json): inputs, calculated metrics, Python references and differences; plus a mesh study and wavelength sweep.
- [`results/g654_spectrum.csv`](results/g654_spectrum.csv): wavelength-dependent computed values for illustrative G.654.E.
- [`tests/test_solver.mjs`](tests/test_solver.mjs): repeatable Node.js numeric tests (`node tests/test_solver.mjs`).
- [`VALIDATION_REPORT.md`](VALIDATION_REPORT.md): result and scope summary.

All snapshot numbers on GitHub reflect the checked-in baseline **only**. Changing interactive inputs does not change repository snapshots until a user commits a new run.

## Related

- [Original Python/Colab FEM code](../Universal_Silica_Fiber_v0_1.ipynb)
- [Original v0.1 description](../README.md)
- [Original numerical verification](../VALIDATION_REPORT.md)
