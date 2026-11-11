# M2 + M3 Upgrade — LS Wave Optics v0.5.0

Updated 2026-11-12. Static GitHub Pages only; no server, login, paid API, external runtime, or Python requirement in the deployed application.

## Code architecture inspected

- `web/solver.mjs`: authoritative scalar radial P1 FEM, pure-silica Sellmeier + fixed Δn default, original fields n_eff/MFD/Aeff/D/S0; **unchanged numerical kernel**.
- `web/app.mjs`: existing presets/analysis/sweeps/import-export; only solver cache-buster changed.
- `web/vector_mode_ui.mjs`: existing Colab JSON visualization (`LSCNS_vector_mode_v1`), preserved.
- `full_vector/vector_modes.py`: existing SciPy sparse coupled Hx/Hy Maxwell FDFD operator, ported to browser with a matrix-free iterative eigensolver.
- `web/validation_*.*` and `web/tests/test_solver.mjs`: M1 validation and prior scalar numerical regression remain intact.
- Original Web COMSOL geometry, materials UI, PML preview, HCF, MCF, PMD, 3D tools untouched.

## Newly added files

| Milestone | File | Function |
|---|---|---|
| M2 v0.4 | `material_models.mjs` | Sellmeier GeO2/F/P2O5, measured table interpolation and complex-index κ storage, material absorption estimate |
| M2 v0.4 | `material_m2_ui.mjs` | Separate material/model lab with scalar step-index FEM evaluation and independent JSON IO |
| M2 v0.4 | `material_m2_ui.css` | M2-only styles |
| M2 v0.4 | `tests/test_material_m2.mjs` | Numeric/reference and default-regression checks |
| M3 v0.5 | `vector_fdfd_browser.mjs` | Real 2D coupled Maxwell Hx/Hy FDFD solver, Ex/Ey/Ez field reconstruction |
| M3 v0.5 | `vector_fdfd_worker.mjs` | Same-origin static module Worker, no UI blocking |
| M3 v0.5 | `vector_m3_ui.mjs` | New Full-vector study panel inside existing Colab viewer |
| M3 v0.5 | `vector_m3_ui.css` | M3-only styles |
| M3 v0.5 | `tests/test_vector_m3.mjs` | Same-grid Python FDFD numeric cross-check |
| M3 v0.5 | `tests/browser_m2_m3_ui.mjs` | Optional Playwright M2/M3 UI interaction regression |

## Minimal changes to original files

```diff
--- a/web/solver.mjs
+++ b/web/solver.mjs
@@
-export const VERSION = '0.3.0';
+export const VERSION = '0.5.0';
```

```diff
--- a/web/app.mjs
+++ b/web/app.mjs
@@
-import {VERSION,PRESETS,claddingIndex,validate,deltaIndex,solve,dispersion,dispersionDiagnostics,spectrum,meshStudy,outputSummary} from './solver.mjs?v=20261112-m1-v03';
+import {VERSION,PRESETS,claddingIndex,validate,deltaIndex,solve,dispersion,dispersionDiagnostics,spectrum,meshStudy,outputSummary} from './solver.mjs?v=20261112-m3-v05';
```

```diff
--- a/web/index.html
+++ b/web/index.html
@@ header
-FIBER DESIGN WORKBENCH · v0.3
+FIBER DESIGN WORKBENCH · v0.5
@@ head, after validation_ui.css
+  <link rel="stylesheet" href="./material_m2_ui.css?v=20261112-m2-v04" />
+  <link rel="stylesheet" href="./vector_m3_ui.css?v=20261112-m3-v05" />
@@ existing app scripts
-  <script type="module" src="./app.mjs?v=20261112-m1-v03"></script>
+  <script type="module" src="./app.mjs?v=20261112-m3-v05"></script>
+  <script type="module" src="./material_m2_ui.mjs?v=20261112-m2-v04"></script>
+  <script type="module" src="./vector_m3_ui.mjs?v=20261112-m3-v05"></script>
```

```diff
--- a/web/vector_mode_ui.mjs
+++ b/web/vector_mode_ui.mjs
@@ old Colab disclaimer
-주의: GitHub Pages는 Full-vector 고유값을 계산하지 않습니다. 이 화면은 Colab 계산 결과를 <b>불러와 표시</b>합니다.
+주의: 위의 기존 Colab 결과 뷰어는 계산하지 않고 <b>JSON을 불러와 표시</b>합니다. 아래 M3 섹션에서 별도 브라우저 벡터 FDFD를 실행할 수 있습니다.
```

No original function name/DOM ID/object schema was removed. M2 uses its own `schemaVersion:1` material JSON while old design JSON remains `{version,model}`. M3 exports existing `LSCNS_vector_mode_v1` field-map format plus additional metadata. The first mode additionally includes `field_map.field_components.{Ex_real,Ey_real,Ez_imag}` with explicit field-phase convention; the old viewer ignores these optional fields. Feature flags set before module loading:
```js
window.LSWaveFeatures={materialM2:false,fullVectorWeb:false};
// Default: both optional features are on.
// Independent M1 flag window.LSWaveFeatures.m1Validation also remains available.
```

## M2 model and assumptions

- Pure silica = ITU-T G Supplement 47, coefficients Malitson; baseline offset `n_core=n_silica+deltaN` unchanged.
- GeO2 = measured fused GeO2 1984 Sellmeier endpoint blended at **coefficient** level with SiO2, limited to ≤30 mol%. The law is an approximation for binary glass.
- F = measured 1%, 2% fluorine-doped silica fits (Applied Optics 22, 3102); simple interpolation 0→1→2% without extrapolation.
- P2O5 = mixture Sellmeier coefficients at ~9.1% from published reference table, interpolate from SiO2 for 0–9.1%.
- `n(λ)` table: columns `wavelength_nm,n_real[,kappa]`, strictly increasing 2–2000 points, linear interpolation only, no extrapolation.
- Complex `n+iκ`: stored and an approximate attenuation `(4πκ/λ)·(10/ln10)·1000` dB/km estimated using a mode intensity core/cladding weighting. It is not a complex eigenvalue or radiation loss calculation.
- Doped Sellmeier validation domain: 400–1700 nm; original solver still supports its old wavelength domain. No claim of dopant-dependent manufacturer material calibration or concentration-dependent absorption.

Primary reference sources:
- https://www.itu.int/epublications/en/publication/itu-t-g-suppl-47-2025-03-general-aspects-of-optical-fibres-and-cables
- https://opg.optica.org/ao/abstract.cfm?URI=ao-23-24-4486
- https://opg.optica.org/ao/abstract.cfm?uri=ao-22-19-3102
- https://www.routledge.com/rsc/downloads/Pages_from_K22108_Book_rev.pdf

## M3 operator and computation

Ported `full_vector/vector_modes.py` lossless, nonmagnetic, z-invariant isotropic transverse magnetic operator:

`β² Hx = (∇t² + k0² εr) Hx - (∂y ln εr) (∂y Hx - ∂x Hy)`

`β² Hy = (∇t² + k0² εr) Hy + (∂x ln εr) (∂y Hx - ∂x Hy)`

Matrix-free 5-point Laplacian and central first derivatives; largest guided eigenvalues via shifted-inverse / diagonal-preconditioned BiCGSTAB; first two x/y-polarized H-field eigenmodes orthogonalized. `Ez` recovered from `curl H` (stored intensity `|Ez|²`); `Ex/Ey` likewise. Neff, MFD second moment, approximate PII, Aeff and polarization energy fractions reported.

**Not implemented**: PML, complex beta, material absorption in the vector eigensystem, bend/leakage loss, HCF vector mode, anisotropic ε, subpixel interface smoothing or arbitrary pixel input. Finite square with Dirichlet boundaries and staircase circle yields grid-dependent neff, so mesh/window convergence is essential.

### Same-grid comparison to existing Python solver

Stored original Python vector FDFD `full_vector/results/vector_mode_result.json` has:
- a = 5.5 µm, λ = 1550 nm, cladding = silica Sellmeier, `Δn=0.0034785793846155`;
- square halfwidth = 22 µm and `grid_n=69`;
- neff (two vector modes) = **1.445711888509771**.

Browser JS FDFD source evaluated on those exact inputs:
- mode 1: **1.4457118884935314**, discrepancy **1.62e-11**;
- mode 2: **1.4457118884934974**, discrepancy **1.63e-11**;
- both eigenpair residuals ≈**9.77e-8**, below the **1e-7** FDFD residual gate;
- Ex/Ey/Ez fractions sum to 1; two dominant polarizations differ.
- These tests check *same-grid algorithmic consistency* and **do not** prove mesh convergence to a continuum Maxwell/COMSOL solution.

## Numerical/regression tests and running

```bash
cd COMSOL/universal_silica_v01/web
node tests/test_solver.mjs
node tests/test_m1_validation.mjs
node tests/test_material_m2.mjs
node tests/test_vector_m3.mjs

# Optional UI + Worker smoke tests; run from the LSCNS repo root
python3 -m http.server 8765
# in another terminal from .../web
node tests/browser_m2_m3_ui.mjs
```

Executed in this session via a JavaScript V8 source evaluator:
- M1 legacy regression **28/28 PASS**;
- M2 pure silica index, material sign trends and table interpolation PASS;
- M2 optional default-offset FEM neff = 1.4461047281745427 vs original = 1.446104728174543 (floating point identical);
- M3 69-grid two-polarization vector solve and comparison to original Python baseline PASS.

The checked-in Node/Playwright test scripts and public website interactive operation were **not run through a local checkout or Playwright in this session**. Validate those independently before production use.

## Scope
Only M2 and M3 added; M4–M6 deferred until explicitly requested. See `CHANGELOG.md`.
