# M1 — Step-index / LP01 / COMSOL public example validation

Version **0.3.0** · 2026-11-12 · Static GitHub Pages · No server or paid APIs

## 1. Original project structure reviewed
- `solver.mjs`: **authoritative existing P1 radial FEM** for LP01, 5-point Gauss quadrature, shifted inverse iteration. Existing outputs: `n_eff`, MFD(PII), Aeff, D, mesh, slope diagnostics, field.
- `app.mjs` and `index.html`: original model, result cards, wavelength sweep, mesh comparison, JSON/CSV/PNG IO. `app.mjs` **only its solver import cache-buster changed; no functions or numerical calculations changed**.
- `../universal_radial.py`, `../Universal_Silica_Fiber_v0_1.ipynb`: Python verification / legacy notebook; **not changed**.
- `mcf_xt.mjs`, `hcf_designer*.mjs`, `pmd*.mjs`, `vector_mode_ui.mjs`: **not changed**.
- `tests/test_solver.mjs`: original 5-preset regression and D/S0/MCF checks; **not changed**.

## 2. New files
| File | Purpose |
|---|---|
| `validation_lp01.mjs` | Independent LP01 Bessel characteristic equation and root |
| `validation_cases.mjs` | COMSOL public example, 2 analytic fixtures and 5 legacy goldens |
| `validation_engine.mjs` | Read-only error/threshold table; 28 numeric tests |
| `validation_ui.mjs` | Feature-flagged validation UI, separate report download |
| `validation_ui.css` | Styles scoped to optional validation section |
| `tests/test_m1_validation.mjs` | Numerical Node regression |
| `tests/browser_m1_validation.mjs` | Optional Playwright UI regression |
| `tests/m1_lp01_scipy_reference.py` | Optional SciPy verification without website Python |
| `CHANGELOG.md` | Version record and validation status |

## 3. Existing-file diffs only

```diff
--- a/web/solver.mjs
+++ b/web/solver.mjs
@@
-export const VERSION = '0.2.2';
+export const VERSION = '0.3.0';
```

```diff
--- a/web/app.mjs
+++ b/web/app.mjs
@@
-import {VERSION,PRESETS,claddingIndex,validate,deltaIndex,solve,dispersion,dispersionDiagnostics,spectrum,meshStudy,outputSummary} from './solver.mjs?v=20261104-comsol-step-v2';
+import {VERSION,PRESETS,claddingIndex,validate,deltaIndex,solve,dispersion,dispersionDiagnostics,spectrum,meshStudy,outputSummary} from './solver.mjs?v=20261112-m1-v03';
```

```diff
--- a/web/index.html
+++ b/web/index.html
@@ head
+  <link rel="stylesheet" href="./validation_ui.css?v=20261112-m1-v03" />
@@ header
-FIBER DESIGN WORKBENCH · v0.2
+FIBER DESIGN WORKBENCH · v0.3
@@ after app module
-  <script type="module" src="./app.mjs?v=20261104-comsol-step-v1"></script>
+  <script type="module" src="./app.mjs?v=20261112-m1-v03"></script>
+  <script type="module" src="./validation_ui.mjs?v=20261112-m1-v03"></script>
```

The new tab is added by the separate JS module to the pre-existing `nav.eng-shortcuts`, immediately after the navigation when selected. No original IDs, functions, result calculations, design schema, or workbench layout were deleted. To turn it off, set `window.LSWaveFeatures = {m1Validation:false}` before the new module loads.

## 4. Public benchmark provenance

COMSOL Step-Index Fiber Application ID 145:
https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html

At free-space 1550 nm: core radius **8 µm**, numerical cladding outer radius **40 µm**, `n_core=1.4457`, `n_clad=1.4378`, published fundamental **neff ≈ 1.4444** (four decimals). Its full-vector HE11 result and the scalar LP01 FEM do **not** have identical governing equations. Its rounded published value cannot establish a high-precision full-vector error. Compare only to a **±0.00005** interval.

Independent scalar LP01 model solves
`u*J1(u)/J0(u) = w*K1(w)/K0(w)`,
`u²+w² = V²`,
`V=k0*a*sqrt(n1²-n2²)`,
`neff = sqrt(n1²-(u/(k0*a))²)`.

For the LP01 fundamental, the first root is between 0 and `min(V,2.4048255577)` even when V>2.405 (the COMSOL reference is multimode at V≈4.895). JS J0/J1 are series; K0/K1 ratio is quadrature; Python uses SciPy's special functions and Brent's method. The analytic model treats the cladding as infinite. The existing FEM has a zero-field outer boundary at 40 µm. This explains finite-mesh and boundary-model differences.

## 5. Inputs and expected values

| Case | a (µm) | R (µm) | ncore | nclad | λ (nm) | expected LP01 analytic neff |
|---|---:|---:|---:|---:|---:|---:|
| COMSOL matching inputs | 8 | 40 | 1.4457 | 1.4378 | 1550 | **1.4444018444092133** |
| Independent single-mode | 4.1 | 40 | 1.4472 | 1.444 | 1550 | **1.444863128995567** |

### Pass/fail rules
| Check | Absolute tolerance |
|---|---:|
| FEM neff vs analytic LP01 | 1e-6 |
| JS analytic vs Python/SciPy stored fixture | 1e-10 |
| FEM neff vs COMSOL rounded `1.4444` | 5e-5 **(rounding only)** |
| Original 5 presets neff vs original numeric snapshot | 1e-8 |
| Original 5 presets MFD(PII), Aeff vs snapshot | 1e-4 µm, 1e-4 µm² |
| Original 5 presets D vs snapshot | 0.005 ps/(nm·km) |
| Eigenproblem algebraic relative residual | 1e-7 |
| Mesh 200/350/550/850 | each finer-mesh analytic error non-increasing, last ≤1e-6 |

### Existing v0.2 snapshot inputs/expected values

All original preset configurations are used **without modification**. Fields below are `neff`, MFD(PII) µm, Aeff µm², D ps/(nm·km).

| Preset | neff | MFD | Aeff | D |
|---|---:|---:|---:|---:|
| g654 | 1.445718967909122 | 12.500052522436938 | 121.92389530338414 | 19.825664345368367 |
| g652 | 1.4461047281745436 | 10.182785332301778 | 79.29336128699885 | 16.99288790683873 |
| graded | 1.4458288150125134 | 10.192939527752662 | 76.88413268997199 | 14.919171294740607 |
| trench | 1.4460435091028094 | 9.839323631520045 | 74.82460399998054 | 17.8038792615834 |
| multi | 1.447001785935375 | 9.582447746912562 | 71.68879956793322 | 19.16257095579606 |

## 6. Run tests

```bash
cd COMSOL/universal_silica_v01/web
node tests/test_solver.mjs
node tests/test_m1_validation.mjs

# Optional independent Python scientific cross-check (Python + scipy)
python tests/m1_lp01_scipy_reference.py

# Optional browser UI smoke (requires Playwright locally, not in GitHub Pages runtime)
python3 -m http.server 8765 --directory /path/to/LSCNS
node tests/browser_m1_validation.mjs
```

For `python -m http.server` use repository root as the server document root. The Node browser test's default URL is `http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/`; this requires running the server from the LSCNS repository root.

## 7. Actual validation performed and limitations

- The **existing repository solver code** and the new analytic/engine source were executed in a JavaScript V8 environment: 28 of 28 numerical comparisons passed.
- COMSOL matching FEM: `1.444401819379318`; LP01: `1.4444018444092133`; |Δ| = `2.5029895e-8`.
- Single-mode FEM: `1.4448631048897878`; LP01: `1.444863128995567`; |Δ| = `2.4105779e-8`.
- Mesh analytic error at 200/350/550/850: `3.51694e-7 / 1.14907e-7 / 4.65406e-8 / 1.94872e-8`.
- **Not yet executed**: Playwright test on the deployed website, or this session's installed Node test script in a checkout.
- **Not validated**: full-vector HE11 fields, PML/complex neff, HCF/MCF structures, birefringence, bend/confinement loss, arbitrary 2D geometries, dopant-dependent n(lambda) or manufacturer data. M2–M6 not implemented.

The validation-report JSON has a separate `schemaVersion:1`; the existing design save/load schema is unchanged, including loading older v0.2 designs.
