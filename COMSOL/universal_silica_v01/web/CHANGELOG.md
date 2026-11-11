# CHANGELOG — LS Wave Optics (Universal Silica Fiber Workbench)

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
- `index.html`: added **two** asset references (M1 JS/CSS) and updated visible UI version text `v0.2 → v0.3`.
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
