# M1 Quantitative Verification — LS Wave Optics v0.5.1

Date: 2026-11-12. Scope is ONLY M1, after M2/M3 v0.5.0. Runs on static GitHub Pages with no server/login/paid API.

## ① Existing repository read / minimal touchpoints

- web/solver.mjs: source of authoritative P1 radial scalar LP01 FEM, neff/MFD/Aeff/D calculations; **only VERSION metadata updated**.
- web/app.mjs: original presets, wavelength sweep, mesh, import/export and live result cards; **only one solver import cache key updated**.
- web/index.html: new M1 asset links, banner version, app/legacy-M1 cache keys; old diagnostic number changed from 06 to 07 for unique headings. No original DOM IDs or layout were removed.
- web/validation_lp01.mjs: the existing independent J0/J1/K0/K1 LP01 characteristic equation, reused **unchanged**.
- web/validation_ui.mjs and validation_engine.mjs: earlier 28-item M1 view **kept**, not replaced or altered in logic. Its historical report version 0.3.0 is retained for compatibility.
- web/material_models.mjs, vector_fdfd_browser.mjs, vector_mode_ui.mjs, HCF, MCF, PMD, geometry, PML and all other tools: **not edited**.

## ② New files

| File | Role |
|---|---|
| m1q_cases.mjs | 4 fixed LP01/SciPy cases and exact v0.5.0 pre-update regression snapshots |
| m1q_engine.mjs | Read-only FEM / LP01 comparison; relative + absolute errors and 5-preset regression |
| m1q_export.mjs | Independent schemaVersion 1 JSON + UTF-8 CSV report writers |
| m1q_ui.mjs | Optional 06 · NUMERICAL VERIFICATION(해석해 대비) with run and download buttons |
| m1q_ui.css | Scoped section styling without affecting existing layouts |
| tests/test_m1_quantitative.mjs | Four-case + five-preset + export Node regression |
| tests/m1q_reference_scipy.py | Separate SciPy Bessel/Brent root cross-check |
| tests/browser_m1_quantitative.mjs | Optional Playwright browser smoke and non-mutation test |
| M1Q_FULL_SOURCE.md | Complete source for all eight new files |

## ③ Full new source code

All eight new source files are copied in full in [M1Q_FULL_SOURCE.md](M1Q_FULL_SOURCE.md).

## ④ Minimal diff against v0.5.0

~~~diff
--- a/web/solver.mjs
+++ b/web/solver.mjs
@@
-export const VERSION = '0.5.0';
+export const VERSION = '0.5.1';
~~~

~~~diff
--- a/web/app.mjs
+++ b/web/app.mjs
@@
-from './solver.mjs?v=20261112-m3-v05';
+from './solver.mjs?v=v0.5.1';
~~~

~~~diff
--- a/web/index.html
+++ b/web/index.html
@@ head
-  <link rel="stylesheet" href="./validation_ui.css?v=20261112-m1-v03" />
+  <link rel="stylesheet" href="./validation_ui.css?v=v0.5.1" />
+  <link rel="stylesheet" href="./m1q_ui.css?v=v0.5.1" />
@@ header
-FIBER DESIGN WORKBENCH · v0.5
+FIBER DESIGN WORKBENCH · v0.5.1
@@ existing diagnostic label only; no ID changes
-06 · SOLVER DIAGNOSTICS
+07 · SOLVER DIAGNOSTICS
@@ original app and historical M1 scripts
-  <script type="module" src="./app.mjs?v=20261112-m3-v05"></script>
-  <script type="module" src="./validation_ui.mjs?v=20261112-m1-v03"></script>
+  <script type="module" src="./app.mjs?v=v0.5.1"></script>
+  <script type="module" src="./validation_ui.mjs?v=v0.5.1"></script>
+  <script type="module" src="./m1q_ui.mjs?v=v0.5.1"></script>
~~~

Official site link for this update:
https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/?v=v0.5.1

Version relationship: banner v0.5.1 = public URL v0.5.1 = new M1 asset queries v0.5.1 = CHANGELOG top section v0.5.1. Historical M2/M3 asset query tags are unchanged because no M2/M3 code was modified. Old M1 report kind/version is intentionally historical; the new verification report uses version=0.5.1, schemaVersion=1.

Disable new feature before module load:
~~~js
window.LSWaveFeatures = {...window.LSWaveFeatures, m1Quantitative: false};
~~~

## ⑤ Benchmark inputs, independent expected neff and tolerances

**IMPORTANT**: the official COMSOL App 145 benchmark is a=8 µm, n_core=1.4457, n_clad=1.4378, 1550 nm. The user-requested a=4.1 µm/Δn=0.005 is NOT an official published COMSOL eigenvalue. It is a separate fixture with explicitly chosen fixed n_clad=1.4440.

| Case | λ nm | a µm | Δn | n_clad (fixed) | R µm | V | independent SciPy LP01 neff | current FEM neff | absolute error | result |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---|
| COMSOL official App 145 geometry | 1550 | 8 | 0.0079 | 1.4378 | 40 | 4.894537730 | 1.4444018444092133 | 1.444401819379318 | 2.502990e-8 | PASS |
| User requested a=4.1 / Δn=0.005 | 1550 | 4.1 | 0.005 | 1.4440 | 40 | 1.998900819 | 1.446081132401304 | 1.4460810799949255 | 5.240638e-8 | PASS |
| Low-V single-mode | 1550 | 4.1 | 0.0032 | 1.4440 | 40 | 1.598623098 | 1.444863128995567 | 1.4448631048897878 | 2.410578e-8 | PASS |
| Other wavelength; fixed index | 1310 | 4.1 | 0.005 | 1.4440 | 40 | 2.365111656 | 1.446608620045317 | 1.4466085591149294 | 6.093039e-8 | PASS |

Each uses 750 radial elements. Case PASS requires: FEM–LP01 |delta neff|<=1e-6, JS Bessel vs independent SciPy |delta neff|<=1e-10, Bessel residual |F(u)|<=1e-8, original FEM algebraic residual<=1e-7 and guided mode. Relative error % = |FEM−LP01|/|LP01|×100 and is displayed but not the primary PASS gate.

Why the FEM tolerance is 1e-6: it allows finite outer boundary and P1 radial spatial discretization errors while remaining much smaller than typical step-index core/cladding index contrast (about 0.003–0.008 here). This is an engineering acceptance criterion, not a COMSOL-certified continuum-error bound.

Bessel formula:
- V = k0·a·sqrt(n_core²−n_clad²)
- u² + w² = V²
- u·J1(u)/J0(u) = w·K1(w)/K0(w)
- neff = sqrt(n_core²−[u/(k0·a)]²)
- fundamental root 0<u<min(V,2.404825557695773); bisection up to 75 iterations or interval width <4e-14 (dimensionless u).
- JS series J0/J1 and K0/K1 quadrature (1024 Simpson intervals) are independent of FEM; Python uses scipy.special and scipy.optimize.brentq (xtol=1e-13).
- A root of the scalar equation is an LP01 approximation, NOT a high-precision vector HE11 result.

COMSOL official source:
https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html

Published COMSOL vector neff≈1.4444 is rounded to four decimal places. Its comparison only checks the ±5×10⁻⁵ rounding interval. It cannot certify vector-mode precision. The user's 4.1 µm condition is COMSOL **미검증**, not incorrectly marked official PASS.

## ⑥ Original function regression before and after

Captured BEFORE v0.5.1 from solver.mjs v0.5.0 blob SHA 6d774b556b83829c77a3b6f3d208eb8ea8da6a3a. Tested again AFTER updating VERSION/import only. Same preset/input, numerical kernels unchanged.

| Preset | neff BEFORE / AFTER | MFD(PII) BEFORE / AFTER µm | Aeff BEFORE / AFTER µm² | D BEFORE / AFTER ps/(nm·km) | Absolute changes |
|---|---|---|---|---|---|
| g654 | 1.4457189679091202 / same | 12.500052522796532 / same | 121.9238953090936 / same | 19.825664231522584 / same | 0 |
| g652 | 1.446104728174543 / same | 10.182785332592127 / same | 79.29336129096376 / same | 16.9928875566912 / same | 0 |
| graded | 1.445828815012516 / same | 10.192939528309598 / same | 76.8841326978277 / same | 14.919171784564464 / same | 0 |
| trench | 1.4460435091028088 / same | 9.839323631718809 / same | 74.82460400259428 / same | 17.80387925775665 / same | 0 |
| multi | 1.447001785935374 / same | 9.582447747061527 / same | 71.68879956968254 / same | 19.162570615215405 / same | 0 |

Regression thresholds: neff<=1e-10, MFD<=1e-6 µm, Aeff<=1e-6 µm², D<=1e-6 ps/(nm·km), each absolute. All 20 before/after metrics passed in the evaluated JS source. Earlier historical M1 28-check code remains loaded separately and unchanged.

## ⑦ How to execute tests

~~~bash
cd COMSOL/universal_silica_v01/web
node tests/test_solver.mjs
node tests/test_m1_validation.mjs
node tests/test_m1_quantitative.mjs

# Optional independent offline scientific check (Python + SciPy)
python tests/m1q_reference_scipy.py

# Browser + download smoke (requires installed Playwright and local HTTP server)
# Run from repository root in terminal 1:
python3 -m http.server 8765
# Run from web/ in terminal 2:
node tests/browser_m1_quantitative.mjs
~~~

Actual validation in this session:
- Source modules executed in a JS evaluation environment: 4/4 LP01 cases PASS, 5/5 presets PASS across 20 neff/MFD/Aeff/D metrics.
- Independent SciPy/Brent root calculations for all 4 cases were executed separately.
- New JS modules parsed and JSON/CSV writers produced valid text: CSV 25 data/header lines (1 header, 4 cases, 20 metric rows).
- Public GitHub Pages page source was updated, but full Playwright browser interaction and Node test files in a cloned checkout were **not executed** in this environment (GitHub CLI network DNS unavailable).
- This scope does not verify material dopants, vector full-wave modes, PML, HCF, MCF, PMD, bend loss, leak loss or measured fiber properties. M2/M3 remain unchanged and are not re-certified by M1 tests.

The new report JSON is independent from saved design JSON. It includes schemaVersion=1 and kind=LSWaveM1QuantitativeVerification; no migration is necessary for saved legacy designs.
