# CHANGELOG — LS Wave Optics (Universal Silica Fiber Workbench)

## v0.7.1 — 2026-11-12 — M2 source audit + M3 true Yee entry (only)

### New files
- `m3_yee_quickstart_v071.mjs`: optional, independently flagged shortcut to the existing **staggered Yee** 2D Maxwell mode solver (`yee_solver.mjs`), plus explicit 1550-nm step-index M3-only preset. The preset disables PML, HCF, bend and absorption and does **not** start a calculation without a second user click.
- `tests/test_m2_m3_acceptance_v071.mjs`: repeatable Node M2 Sellmeier, α↔κ, CSV interpolation, M1 regression and real Yee N24 test with measured process time/RSS.
- `M2_M3_V071_ACCEPTANCE.md` / `M2_M3_V071_FULL_SOURCE.md`: verified inputs, references, comparison/mesh data, minimal diffs, full new code and open acceptance gaps.

### Minimal old-file changes
- `m2_preview_engine_v06.mjs`: corrected pure fused-silica primary citation to I.H. Malitson (1965), JOSA 55, 1205, DOI 10.1364/JOSA.55.001205; **zero coefficient/value changes**.
- `m2_preview_ui_v06.mjs`: separate Malitson SiO₂ vs Fleming GeO₂ literature links and refresh module cache key; no change to n(λ), κ(λ), tables, interpolation or save schema.
- `solver.mjs`: VERSION `0.7.0 → 0.7.1` only (no changes to FEM numerical logic).
- `app.mjs`: only the solver import URL cache key changed to `?v=v0.7.1`.
- `index.html`: header version to `v0.7.1`, app/M2 cache keys to `v0.7.1`, one additional M3 shortcut module script reference. Public URL `?v=v0.7.1`.
- Existing collocated M3 v0.7 study, true Yee Maxwell solver, original Colab JSON viewer, old scalar solver, M1, M2 value model, M4–M6 code, DOM IDs and saved-design JSON are **unchanged**.

### What is validated
- Current scalar M1: 4/4 analytic cases and 5/5 old presets passed without changing their numeric values; original 28-item M1 suite remains in place.
- M2 at 1550 nm: pure silica = 1.4440236217032607 (exact original Sellmeier), GeO₂ 5 mol% modeled = 1.4515267771427718, F 1 **wt%** fit = 1.4393958509698328, P₂O₅ 5 mol% modeled = 1.452224423108019. F mol% still **requires user-sourced coefficients**.
- M2 index table supports linear/PCHIP with no extrapolation, complex `n+iκ` and α[dB/km]↔κ conversions. M2 values and references were re-evaluated from checked-in source in this session.
- Strict M3 engine: existing `yee_solver.mjs` evaluated directly on a 24×24 Yee grid, two polarizations, neff≈1.4460854785569 / 1.4460854785554, eigenpair residuals 4.73×10⁻⁹ / 4.82×10⁻⁹, portable schema `LSCNS_vector_mode_v1`. This is a coarse-grid numerical test, **not** a claim of mesh convergence.
- See `M3_M6_YEE_VALIDATION_20261112.md` for existing source-backed N24/48/72/96 neff, M1/scalar/Colab differences and measured Node process RSS/time. Last-grid convergence criterion is **NOT met**, and that remains explicitly unresolved.
- Public browser click-through of the new shortcut, published-site Worker and user-machine memory measurement **not executed** in this session. The Node test is checked in, but has not been executed in a cloned checkout.

### Strict scope / limits
No M4 PML, HCF, bending loss, M5 geometry, M6 workflows, or underlying Yee Maxwell operator was changed. The new quickstart is UI navigation, not a new physics algorithm.

## 2026-11-12 · M3–M6 experimental Yee workbench

- Added a separate browser Yee full-vector solver with complex ILU(0)-GMRES, block shift-invert/Ritz, Worker progress/cancellation, complex fields, polarization, MFD/Aeff, and Colab-compatible vector JSON.
- Added complex-coordinate PML, optional M2 absorption, conformal equivalent-index bends, HCF geometry sampling, candidate-mode filters and explicit loss-convergence status.
- Added polygon/SVG/image-mask input, material assignments, subpixel sampling and direct N-core supermodes with reduced coherent coupling indicators; retained existing interpolation and scalar solvers.
- Added workflow links to existing panels, 1–2 parameter sweeps, CSV, saved overlays and project schema3 adapters for existing Parameters/CAD/Materials/PML/M2/scalar state.
- Acceptance is incomplete: mesh/PML loss convergence and quantitative HCF loss remain unverified. Browser acceptance passed: project-file roundtrip with scalar/CAD/parameter restoration, 2-axis sweep CSV, cancellation, circle geometry, legacy import and layout switching. See M3_M6_YEE_VALIDATION_20261112.md for measured tables and limits.


## v0.7.0 — 2026-11-12 — M3 vector FDFD study controls and comparisons

### Added (new files, optional)
- `m3_study_metrics_v07.mjs`: bounded odd grid N=29..91; conditional same-grid Colab comparisons; independent M1 LP01 and scalar FEM references, non-monotone mesh-change indicators, relative eigen-residual gates; conservative Float64 allocation model distinctly labeled NOT measured memory.
- `m3_study_worker_v07.mjs`: isolated module Web Worker, measured wall-clock via `performance.now()`, per-grid/inner-eigenpair progress, two-mode Ex/Ey/Ez maps exported in unchanged `LSCNS_vector_mode_v1` schema; user cancel uses `worker.terminate()`.
- `m3_study_ui_v07.mjs` and scoped CSS: separate M3 study section inside existing Full-vector tool, five-or-fewer grid convergence table, LP01/scalar/Colab difference, per-grid time/memory status, explicit estimated-time warning, selectable polarizations/maps, matched-grid Colab overlay, cancel, portable vector JSON and separately versioned `LSWaveFullVectorStudy` report.
- `tests/test_m3_study_v07.mjs`, `tests/benchmark_m3_v07.mjs`, `tests/browser_m2_m3_v07.mjs`: numeric compare, local Node RSS/heap benchmark and optional Playwright UI regression.
- `M2_M3_V07_VALIDATION.md`, `M2_M3_V07_FULL_SOURCE.md`: complete code, numerical comparisons, regression checks, minimal old-file diff, assumptions.

### Existing-file changes (minimal)
- `index.html`: M2/M3 CSS+JS module links, displayed header and app import cache key `v0.7.0`.
- `solver.mjs`: `VERSION` string only `0.5.1 → 0.7.0`. No changes to P1 FEM, n(λ), D, S0, MFD, Aeff.
- `app.mjs`: only one `solver.mjs` import cache key changed. Original DOM IDs, functions, saved-design JSON and Colab viewer untouched.
- Old `vector_fdfd_browser.mjs` FDFD eigen operator, `vector_m3_ui.mjs`, `vector_mode_ui.mjs` left unchanged; new M3 layer adds workflow, not a replacement algorithm.
- Flag: `window.LSWaveFeatures.fullVectorStudyV07=false` before module load hides only the new study.

### Measured and unmeasured
- Evaluated JS source on a V8 host, same input as stored Python 69×69: Python neff=1.445711888509771 vs JS 1.445711888493530, Δ=1.6241×10⁻¹¹; algebraic eigenpair residual=9.774×10⁻⁸.
- M1 analytic LP01 neff=1.445718990544404; original scalar FEM=1.4457189679091202. Browser FDFD 69-grid diff from LP01=7.1021×10⁻⁶, from scalar=7.0794×10⁻⁶.
- Reference V8 wall times (two modes): N 29: 1.276 s; 41: 4.088 s; 51: 6.611 s; 69: 14.816 s; 81: 22.835 s (not measured on the user's browser).
- Grid neff sequence non-monotone due staircased interface; convergence status is **미수렴** if adjacent |Δneff| exceeds 1×10⁻⁶. No continuum validation claim.
- **No memory measurement** could be obtained in the current V8 execution environment. Browser Worker attempts performance.memory if available and otherwise prints `실측 미지원`; typed-array size is an explicitly separated allocation model, not RAM. Node benchmark script can measure RSS/heap on the user's own machine.
- Browser/Playwright click/download/cancel tests were created but NOT executed against public GitHub Pages during this session.

### Explicit M3 acceptance gap
- Existing solver is coupled transverse-Hx/Hy **collocated** FDFD with matrix-free shifted-inverse BiCGSTAB (inner linear tolerance 2×10⁻⁹, relative eigenpair tolerance 1×10⁻⁷).
- It is **NOT a Yee-staggered grid and NOT an edge-element vector FEM**. Therefore the exact user request "Yee FDFD or edge/vector FEM" is **not fully satisfied**. The source and UI label this honestly; no unsupported discretization is claimed.
- PML, complex effective index, confinement/bend losses, HCF and material absorption in eigenproblem remain out of scope (M4).

## v0.6.0 — 2026-11-12 — M2 independent material preview

### Added
- `m2_preview_engine_v06.mjs`: exact legacy SiO₂+Δn, sourced GeO₂ end-member Sellmeier mix, P₂O₅ 9.1 mol% anchor, explicitly **1/2 wt%** fluorine measured-fit interpolation, user-coefficient-only F mol% (NO ungrounded concentration conversion), CSV n(λ) linear/PCHIP interpolation with no out-of-range extrapolation, n+iκ storage and κ↔α conversion.
- `m2_preview_ui_v06.mjs`, scoped CSS: separate Core/Cladding material selector and concentration forms, model provenance, dual n/k wavelength charts, reference curves and 1550 nm table, distinct JSON export/import (schemaVersion:2 with cautious migration of old ambiguous F schemaVersion:1).
- `tests/test_m2_preview_v06.mjs`: coefficient reference values, no-F-mol%-guess guard, interpolation, α↔κ round trip, legacy silica and FEM non-regression.
- Default scalar Radial FEM **not connected to new material preview**. Original core Δn offset remains default and identical; M2 schema extension does not change saved-design JSON.
- `material_m2_ui.mjs`: only two labels corrected (previous F % incorrectly called mol%, actual original Fleming/Wood specimens wt%). No numerical changes.

### Verified n(1550 nm), V8 source execution
- Pure SiO₂: 1.4440236217032607, identical to original `silicaIndex(1550)`.
- GeO₂ 5 mol% modeled: 1.4515267771427718.
- F-doped silica 1 **wt%** empirical fit: 1.4393958509698328 (**not 1 mol%**).
- P₂O₅ 5 mol% interpolated: 1.452224423108019.
- At λ=1550 nm, α=0.2 dB/km yields κ=5.6802508913155675×10⁻¹²; inverse produces 0.2 dB/km.
- Dopant concentration mixing is a modeling approximation, NOT empirical index verification at every concentration.

### Sources (coefficients)
- SiO₂/GeO₂: Fleming, Applied Optics 23, 4486 (1984), Table II, https://doi.org/10.1364/AO.23.004486
- F doped 1/2 wt%: Fleming & Wood, Applied Optics 22, 3102 (1983), Table II, https://doi.org/10.1364/AO.22.003102
- P₂O₅ 9.1 mol%: Optical Fiber Communication Systems with MATLAB and Simulink Models, Table 3.4; https://www.routledge.com/rsc/downloads/Pages_from_K22108_Book_rev.pdf

### Preserved
M1 v0.5.1: 4/4 LP01 cases and 5/5 legacy preset regressions; older M1 28-check method unchanged. No M4, M5, M6 changes.

## v0.5.1 — 2026-11-12 — M1 Quantitative Verification only

### Added (independent, feature-flagged)
- `m1q_cases.mjs`: four fixed step-index LP01 fixtures (actual official COMSOL App 145: a=8 µm; separate user-requested a=4.1 µm/Δn=0.005; V<2.405; λ=1310 nm) and exact v0.5.0 pre-update five-preset numeric snapshots.
- `m1q_engine.mjs`: existing scalar Radial FEM compared read-only to independent `validation_lp01.mjs` Bessel oracle with independently computed SciPy reference numbers; before/after neff/MFD/Aeff/D regression for five original presets.
- `m1q_export.mjs`: separate schemaVersion:1 M1 verification JSON and UTF-8 CSV; existing design/result JSON/CSV/PNG unchanged.
- `m1q_ui.mjs` / `m1q_ui.css`: optional `06 · NUMERICAL VERIFICATION(해석해 대비)` section, four-case absolute/relative error table, PASS/FAIL, old/new comparison table and Export buttons. Disable by setting `window.LSWaveFeatures.m1Quantitative=false` before module load.
- `tests/test_m1_quantitative.mjs`, `tests/m1q_reference_scipy.py`, `tests/browser_m1_quantitative.mjs`: fixed-case, SciPy Bessel/Brent, previous-output preservation, export and optional Playwright browser tests.
- `M1Q_VALIDATION.md` and `M1Q_FULL_SOURCE.md`: complete source code, minimal diffs, benchmark provenance, error tolerances and verification caveats.

### Minimal existing file patches
- `solver.mjs`: **VERSION only** 0.5.0 → 0.5.1 (no changes to FEM assembly/dispersion).
- `app.mjs`: original `solver.mjs` import cache URL only updated to `?v=v0.5.1`.
- `index.html`: header `v0.5 → v0.5.1`, app and existing legacy M1 static asset URLs `?v=v0.5.1`, plus the new M1 CSS/JS import. Static `06 · SOLVER DIAGNOSTICS` display number updated to 07 to avoid collision; existing diagnostics ID/logic unchanged.
- No modifications to M2/M3, HCF, MCF, PMD, Full-vector, Geometry, PML, saved design schema, original controls or user data.
- **Version alignment**: page banner `v0.5.1` = public URL `?v=v0.5.1` = M1 asset cache `?v=v0.5.1` = CHANGELOG `v0.5.1`. Older module-specific asset tags retain their original historical versions because their code was not changed.

### Benchmarks and validation — JS source executed
- Official COMSOL App 145 source parameters: 1550 nm, a=8 µm, ncore=1.4457, nclad=1.4378; public rounded vector HE11 neff≈1.4444. Independent scalar LP01 analytic neff=1.4444018444092133; existing P1 FEM=1.444401819379318; |Δ|=2.50299e-8.
- **User-requested separate** a=4.1 µm, Δn=0.005, fixed nclad=1.4440 at 1550 nm: LP01=1.446081132401304; FEM=1.4460810799949255; |Δ|=5.24064e-8. This is **not a published COMSOL reference value**.
- Low V case (V≈1.598623, Δn=0.0032): LP01=1.444863128995567; FEM=1.4448631048897878; |Δ|=2.41058e-8.
- 1310 nm, a=4.1 µm, Δn=0.005, fixed nclad=1.4440: LP01=1.446608620045317; FEM=1.4466085591149294; |Δ|=6.09304e-8. Actual material dispersion is **unverified**.
- **4/4** scalar FEM–LP01 cases PASS with |Δneff|≤1e-6; independent SciPy reference comparison tolerance ≤1e-10 and FEM eigen-residual ≤1e-7.
- Five v0.5.0 presets have **identical current neff, MFD(PII), Aeff and D** (20/20 before/after metrics within numerical tolerance). The previous M1 28-check numerical regression remains available.
- Numerical engine and CSV/JSON serialization were executed as source in the JS evaluation environment; independent SciPy Bessel roots were evaluated. The checked-in optional Node/Playwright tests and deployed-browser click operations have **not** yet been executed against the public site.

### Explicit exclusions
Not validated: user's 4.1 µm model against actual COMSOL output, full-vector/HE11 precision outside published four-digit result, n(λ) dispersion at fixed nclad, M2 material physical accuracy, M3 vector convergence, PML, HCF, MCF, bend/confinement losses, manufacturer/fiber measurements. No M2–M6 work performed in this patch.

## v0.5.0 — 2026-11-12 — M3 Browser full-vector FDFD

### Added
- `vector_fdfd_browser.mjs`: real 2D coupled transverse-H Maxwell FDFD eigensolver (Hx, Hy), matrix-free shifted-inverse BiCGSTAB, reconstructed Ex/Ey/Ez and modal metrics (neff, MFD, Aeff, polarization). Raw Ex/Ey real and Ez imaginary amplitudes are also exported; legacy intensity-map fields remain intact.
- `vector_fdfd_worker.mjs`: runs full-vector calculations off the UI thread using a module Web Worker. Static GitHub Pages compatible; no API, account, paid compute or server.
- `vector_m3_ui.mjs` / `vector_m3_ui.css`: separate controls and vector-field maps **inside the existing Full-vector study**. The existing Colab JSON reference viewer is retained unchanged in functionality.
- `tests/test_vector_m3.mjs` / `tests/browser_m2_m3_ui.mjs`: same-grid Python numeric reference comparison and optional browser UI check. JSON exported in the pre-existing `LSCNS_vector_mode_v1` format.
- `M2_M3_VALIDATION.md` and `M2_M3_FULL_SOURCE.md`: full source listing for 10 new files, minimal existing-file diffs, independent references, executable tests, and explicit limitations.

### Minimal existing-file changes
- `index.html`: new M2/M3 CSS and JS imports, visible version to v0.5, cache-bust original `app.mjs`.
- `solver.mjs`: only VERSION `0.3.0 → 0.5.0` (metadata); no modification to scalar FEM assembly or fields.
- `app.mjs`: only `solver.mjs` import URL cache key updated, no calculation function changes.
- `vector_mode_ui.mjs`: one disclaimer sentence clarified so the original viewer is no longer described as the only full-vector capability.

### Numeric checks executed
- Same physical configuration/discretization as existing `full_vector/results/vector_mode_result.json`: λ=1550 nm, a=5.5 µm, delta_n=0.0034785793846155, silica Sellmeier cladding, 69×69 grid, square window halfwidth=22 µm.
- Existing Python SciPy FDFD reference: neff=1.445711888509771. JS FDFD two polarizations after tighter inverse iteration: 1.4457118884935314 and 1.4457118884934974; absolute discrepancies ~1.62e-11 and 1.63e-11. This is a same-grid solver cross-check, not COMSOL certified HE11 accuracy.
- Both eigen residuals <1e-7, Ex/Ey/Ez nonnegative intensities with unit total energy fraction, modes have orthogonal dominant polarization. Grid-index convergence is not guaranteed without interface smoothing.
- Old M1 baseline `28/28` numerical checks continued to pass.
- Browser Playwright click test has been written but **not executed against the public GitHub Pages site in this session**.

### Limitations
Real isotropic epsilon, no PML or radiation/leakage/bend losses; rectangular finite difference grid with staircased circle, Dirichlet boundary and approximate reconstructed Ez. Complex M2 κ is not used in the M3 eigenproblem. This is a genuine coupled Maxwell FDFD prototype with an eigenpair residual, *not* a validated general-purpose full-vector COMSOL replacement.

## v0.4.0 — 2026-11-12 — M2 Material model extension

### Added
- `material_models.mjs`: optional real Sellmeier material families for SiO2–GeO2, SiO2–F (measured 1/2%), SiO2–P2O5 (9.1% reference), fixed Δn offsets, strictly sorted experimental n(λ) table and user-entered κ.
- `material_m2_ui.mjs` / `material_m2_ui.css`: separate optional Material Research section, core/cladding indices and isolated scalar step-index solve, material JSON import/export with schemaVersion 1.
- `tests/test_material_m2.mjs`: coefficients, concentration boundaries, table range/interpolation, and unchanged legacy default neff.
- Complex n is stored as (n, κ) and supports a perturbative weak-loss absorption estimate, not a computed complex eigenmode/PML loss.

### Regression and compatibility
- Original user-facing model remains pure-silica Sellmeier + Δn by default; no modifications to `solve()`, `dispersion()`, `spectrum()` or existing design save/load.
- Baseline G.652.D example: old solver neff=1.446104728174543; optional M2 default composition neff=1.4461047281745427 (floating-point equivalent).
- F 1% n(1550)≈1.439396, F 2%≈1.435212, GeO2 6.3% interpolation≈1.453471, P2O5 9.1% reference≈1.458942. These are coefficient-model results, not manufacturer fiber measurements.
- No arbitrary concentration extrapolation: GeO2 ≤30%, F ≤2%, P2O5 ≤9.1%. Doped model limited to 400–1700 nm. Baseline broad-wavelength legacy solver untouched.

### Sources
- SiO2: ITU-T G Suppl. 47 Table I.1, 2025.
- GeO2: Fleming, Applied Optics 23, 4486 (1984).
- F: Applied Optics 22, 3102 (1983), measured 1% and 2%.
- P2O5: Optical Fibers: Signal Attenuation and Dispersion 9.1% mixture coefficient table (secondary reproduction).
- FDFD physics: existing repo `full_vector/vector_modes.py`.

## v0.3.0 — 2026-11-12 — M1 Validation only

### Added (new files)
- `validation_lp01.mjs`: independent lossless step-index scalar LP01 Bessel characteristic-equation root solver, with series J0/J1 and quadrature K1/K0.
- `validation_cases.mjs`: authoritative COMSOL step-index parameter fixture (App ID 145), independent SciPy single-mode benchmark and unchanged v0.2 five-profile regression goldens.
- `validation_engine.mjs`: deterministic checks for FEM vs LP01, published COMSOL rounded neff, algebraic residual, mesh-convergence trend and original neff/MFD/Aeff/D regressions (28 total).
- `validation_ui.mjs` and `validation_ui.css`: independently mounted M1 validation section, error table, exact tolerance display and separately downloadable versioned validation-report JSON. Feature flag: set `window.LSWaveFeatures = {m1Validation:false}` before module load to disable.
- `tests/test_m1_validation.mjs`: 28/28 numeric assertions plus independent LP01 root/mesh guards.
- `tests/browser_m1_validation.mjs`: optional Playwright UI smoke including non-mutation of active preset, original neff cards and the existing controls.
- `tests/m1_lp01_scipy_reference.py`: optional independent Python/SciPy root reference cross-check.
- `M1_VALIDATION.md`: benchmark provenance, method, assumptions, tolerances and commands.
- `M1_FULL_SOURCE.md`: copy-ready full source listing for all eight new modules, assets and regression scripts.

### Changed (minimum connection)
- `solver.mjs`: **only** `VERSION` metadata `0.2.2 → 0.3.0`; solver assembly, convergence, material index, dispersion and result shape are unchanged.
- `index.html`: added **two** asset references (M1 JS/CSS), updated visible UI version text `v0.2 → v0.3`, and bumped the app script cache key.
- `app.mjs`: changed **only** the `solver.mjs` import cache key; no numerical functions or behavior changed.
- No existing DOM ID/function/JSON design object was renamed or removed. Design JSON continues to use `{version,model}`. The new *validation report* has `schemaVersion:1`, separate from saved design JSON. No migration is required for existing designs.

### Validation status
- Executed JavaScript FEM and independent Bessel oracle directly on 2026-11-12:
  - COMSOL benchmark geometry: scalar analytic neff 1.4444018444092133; existing FEM (750 elements) 1.444401819379318; difference 2.503e-8.
  - Published COMSOL reference: 1.4444 **rounded to four decimal places**; FEM is within ±0.00005. This is not a full-vector accuracy certification.
  - Additional single-mode fixture: analytic 1.444863128995567; FEM 1.4448631048897878; difference 2.411e-8.
  - All **28/28** defined numeric checks passed in executed V8 source-module test, including five legacy preset profiles.
  - Mesh benchmark |Δneff|: 200→3.517e-7, 350→1.149e-7, 550→4.654e-8, 850→1.949e-8.
- Browser click/Playwright regression and the command-line Node/Python test scripts were written, but **not executed on the published website in this session**.

### Explicitly out of scope
M2–M6; new refractive-index laws, full-vector browser FEM/FDFD, PML solver, confinement/bend loss, HCF solver, MCF N-core solver, material absorption and arbitrary 2D mask meshing. M1 establishes numerical comparators only.
