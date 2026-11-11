# CHANGELOG — LS Wave Optics (Universal Silica Fiber Workbench)

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
