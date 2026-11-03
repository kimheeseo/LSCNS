# Universal Silica Fiber Workbench — Web UI v0.2.1

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
- **Not supported by the LP01 physical FEM:** arbitrary 2D cross sections, microstructured PCF, ARF/NANF HCF, full-vector electromagnetic modes, polarization, birefringence, material attenuation from geometry alone, bending radiation loss from complex modes/PML, confinement loss, or model inversion of unknown manufacturer profiles. **Separate** MCF interpolation, measured-anchor attenuation/bend estimation, and statistical PMD modules are documented below.
- **Domain:** model calculator radius is a numerical boundary, **not** actual 125 µm fiber cladding. Edge-truncation and mesh convergence should be checked for research use. Small solver residual does not establish physical accuracy.
- A G.654.E-like preset uses **assumed 5.5 µm core radius and fitted Δn** to target MFD ≈12.5 µm. This is not independent LE125 certification or COMSOL license-equivalent validation.

## Dispersion-slope diagnostics (2026-11-03)

The web UI now displays three **separate calculated quantities** directly beneath the four original mode metrics, after **모드 계산 실행**:

- **S₁₅₅₀**, the dispersion slope at 1550 nm: \([D(1560)-D(1540)]/20\), units ps/(nm²·km).
- **λ₀**, the zero-dispersion wavelength where \(D(λ₀)=0\), located by scanning 1100–1550 nm and a bracketed root search.
- **S₀**, the dispersion slope **at λ₀**, not at 1550 nm.

All dispersion values use five-point finite differences of freshly calculated radial-FEM `neff(λ)`. The output is recomputed after model changes and stored in exported JSON/CSV. If no zero crossing is found in the search window, λ₀ and S₀ appear as **not available** rather than guessed.

**G.654.E-like example** (`a=5.5 µm`, `Δn≈0.003478579`, 750 elements): S₁₅₅₀ ≈ **0.059346 ps/(nm²·km)**, λ₀ ≈ **1278.901 nm**, S₀ ≈ **0.093321 ps/(nm²·km)**. These are **fitted surrogate-model calculations**, not measured PremierLink LE125 specifications. Always verify whether a manufacturer document refers to `S₀` or `S₁₅₅₀` before comparing.

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

## New engineering-analysis modules — attenuation → macrobend → PMD (2026-11-03)

All three modules have been added to the **same GitHub Pages workbench**, beneath the original LP₀₁ mode analysis. At the top of the results area, use the **01 감쇠 / 링크 버짓**, **02 굽힘손실**, **02+ 굽힘 물리식 / COMSOL 후처리** and **03 PMD / DGD** shortcuts to jump directly to the sections.

- **Phase 1 — attenuation**: enter measured/specification α1550 and α1625, interpolate only for 1550–1625 nm, combine fiber length, connector and splice losses, compute output optical power and plot length-vs-loss. Initial numbers are illustrative, not real YOFC/LE125 measurements.
- **Phase 2 — macrobending**: input four measured dB/turn anchors at two bending radii and 1550/1625 nm; interpolate **only** between measured radii/wavelengths, combine turn count with an optional baseline from Phase 1 and draw the radius-vs-loss curve. This is **not** a complex-neff curved-fiber FEM/PML model.
- **Phase 3 — PMD**: use input PMD coefficient and length in a **strong random coupling / Maxwell DGD statistical** model. Explicitly select whether the input coefficient refers to mean or RMS DGD. Output mean, RMS, median/P95/P99 DGD and statistical density plot. It **does not** compute actual vector-polarized fiber modes, Jones matrices, or standardized cable PMD_Q.

Each section exports JSON results with model provenance and limitations. Existing LP₀₁ solver, MCF illustration, YOFC and LE125 settings remain available.

**Full instructions, limitations, formulas and verification examples:** [ENGINEERING_MODULES.md](./ENGINEERING_MODULES.md).

**New regression tests:**
```sh
node COMSOL/universal_silica_v01/web/tests/test_attenuation.mjs
node COMSOL/universal_silica_v01/web/tests/test_bending_loss.mjs
node COMSOL/universal_silica_v01/web/tests/test_pmd.mjs
```

**Fidelity caveat:** the phrase "prediction" for these new outputs is limited to **measurement-input-driven extrapolation-free interpolation or theoretical statistics**. The current scalar LP₀₁ FEM still does not independently predict material attenuation, radiation macro-bending leakage or actual fiber PMD from geometrical input alone.

## Velamuri 2018 bending paper extension — conformal mapping + loss postprocessing

The workbench now has a separate **02+ · 굽힘 물리식 / COMSOL 후처리** section, implemented in [bend_physics_bridge.mjs](bend_physics_bridge.mjs) with a dedicated [browser UI](bend_physics_ui.mjs). It references Ashitosh Velamuri et al., 2018, [*Study of Bending Losses in Optical Fibers using COMSOL*](https://www.comsol.com/paper/study-of-bending-losses-in-optical-fibers-using-comsol-multiphysics-61111).

- Visualize `n_equiv(x)=n_stress(x)(1+x/R)` for the currently selected step/graded/trench/multilayer index design; stress OFF or **simple beam/Poisson proxy (not GEBT)**.
- Convert **externally computed** complex mode `|Im(n_eff)|` to dB/m and dB/turn (bend **diameter**, not radius, is the UI input).
- Verify arithmetic via the paper's 1550 nm, **10 mm bend diameter**, BIF simulated `0.0119 dB/turn` example: the inverse imaginary mode index `1.0758076e-8` gives the same loss. **This is mathematical inversion, not a new FEM prediction.**
- Provide optional diagnostic for **user-supplied** PML thickness / imaginary mode index data; display the paper's 7λ example as a physical thickness. Neither PML nor a complex 2D mode solver is implemented in the browser.

The prior measured-data `bending_loss.mjs` module is retained. Detailed math, caveats, tests and paper reference are in [ENGINEERING_MODULES.md](ENGINEERING_MODULES.md), with a numerical [Velamuri paper implementation and validation note](VELAMURI_PAPER_VALIDATION.md).

```sh
node COMSOL/universal_silica_v01/web/tests/test_bend_physics.mjs
```

## Paper-based 2D scalar FEM/PML mode solver (2026-11-03)

This extends the earlier Velamuri complex-index *postprocessor* with a separate **actual 2D scalar P1 triangular FEM + complex-coordinate PML eigensolver** written in Python/SciPy. It **computes its own complex eigenvalue** and the corresponding 2D intensity distribution (not just interpolating measured loss or receiving Im(n_eff)).

- [**2D Python FEM source**](../bending_fem/solver.py)
- [**Colab simulation notebook**](../bending_fem/Velamuri_Bend_2D_PML.ipynb)
- [**Executed numerical record / not converged**](../bending_fem/example_validation.json)
- [**Detailed validation report**](../bending_fem/VALIDATION_REPORT.md)
- [**Full instructions and limits**](../bending_fem/README.md)

The web workbench now includes shortcut **02++ · 실제 2D FEM/PML 결과**. Run the Colab notebook and import `bend2d_pml_result.json`. The **browser does not execute the SciPy eigenvalue solve**; it displays imported numerical results and mesh/PML sensitivity, warning if numerical convergence criteria are not met.

**Actual illustrative run:** G.652-like hypothetical step index, λ1550, core radius 4.1µm, Δn=.005, bend **diameter** 10mm: Re(neff)=1.446327526, Im(neff)=6.1073e−5. It converts to **67.5562 dB/turn**, but **FAILED numerical convergence screening** (mesh last-two change 4.729%, PML 6λ→7λ change 3.507%, both > user-selected 2%). This value is **NOT a reliable or certified fiber bend-loss prediction** and is **NOT the paper's BIF geometry**.

The research solver is **scalar, not full-vector**; GEBT mechanics, physical 125µm cladding/coating, actual manufacturer RIP, and measured product verification are not implemented. The earlier 1D LP01 web solver and separate measured-anchor bending module remain unchanged.

```sh
pip install numpy scipy matplotlib pandas pytest
python -m pytest COMSOL/universal_silica_v01/bending_fem/tests/test_solver.py -q
```

## MCF · Ideal two-core Cross-talk (2026-11-03)

Select the **MCF · 2코어 XT 분석** preset/profile in the existing web UI.
Inputs: two-core pitch 12–24 µm and propagation length 0–100 mm.
Outputs: even/odd supermode Δneff, κ, complete-transfer length Lc,
ideal coherent core 1/2 fractions, and **XT = 10 log10(P2/P1)** in dB.
The cross-talk curve is drawn against propagation length and included in
downloadable CSV/JSON; graph values are visually clipped to [-60, 40] dB,
while the exact transfer null/full-transfer results display −∞/+∞.

**Interpretation:** the browser does NOT calculate a new 2D FEM solution.
It log-interpolates four previously archived legacy M8 scalar-FEM supermode
splittings (pitch 12/16/20/24 µm) at λ=1550 nm, a=4.1 µm and Δn=0.005.
The resulting oscillatory coupling is a lossless, identical-core coherent
model. It cannot predict real **MCF cable XT (dB/km)**, random long-distance
coupling, bend/twist, wavelength-dependent XT, polarization, or measured
product performance. The improved M8 conforming-FEM validation is distinct
from these archived legacy pitch references.

Numerical source: [mcf_xt.mjs](mcf_xt.mjs). Tests:
`node COMSOL/universal_silica_v01/web/tests/test_solver.mjs`.
