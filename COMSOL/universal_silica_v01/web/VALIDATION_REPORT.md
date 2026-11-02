# Browser FEM v0.2 — Numerical validation and UI smoke check

Date: 2026-11-03 (KST)

## 1. Python vs JavaScript numerical cross-check

Reference engine is the existing `universal_radial.py` from v0.1. Both use the same material assumptions, 750-element radial mesh, 5-point Gauss quadrature, scalar P1 FEM, numerical waveguide eigenmode and 5-point dispersion. The browser uses deterministic shifted inverse iteration instead of SciPy sparse ARPACK.

| Preset | JS neff | JS MFD PII (µm) | JS Aeff (µm²) | JS D (ps/nm/km) |
|---|---:|---:|---:|---:|
| g654 | 1.445718968 | 12.50005 | 121.92390 | 19.82566 |
| g652 | 1.446104728 | 10.18279 | 79.29336 | 16.99289 |
| graded | 1.445828815 | 10.19294 | 76.88413 | 14.91917 |
| trench | 1.446043509 | 9.83932 | 74.82460 | 17.80388 |
| multi | 1.447001786 | 9.58245 | 71.68880 | 19.16257 |

Maximum absolute difference to the independent Python implementation:

- neff: **2.665e-15**
- MFD Petermann-II: **5.569e-10 µm**
- Aeff: **7.856e-09 µm²**
- D: **4.898e-07 ps/(nm·km)**

All 5 presets pass the checked numeric tolerances (neff 1e-8; MFD 1e-4 µm; Aeff 1e-4 µm²; D 0.005 ps/(nm·km)). The Node `tests/test_solver.mjs` additionally checks residual, basic guidance, mesh trend, spectrum and invalid-model guards. These are numerical *consistency* checks, not manufacturer or experimental validation.

## 2. Runtime / interface verification

- Browser UI rendering and interactive click test run using Chromium Playwright with locally injected HTML/CSS/ES modules (host sandbox blocked direct local navigation).
- Confirmed: baseline G.654.E values rendered, switching to Graded-index invalidates cards, clicking **모드 계산 실행** recomputes results, **메시 수렴** populates 4 rows, no page JS exceptions.
- The deployed GitHub Pages URL is version-controlled but external hosting/caching must be verified separately; no claim of automated, live GitHub commit on parameter change.

## 3. Limitations

Only axisymmetric scalar LP01 weak-guidance mode. The model omits full-vector Maxwell, stress, dopant-dependent silica Sellmeier, true cladding and bending/radiation loss, PML, arbitrary 2D microstructure, and real product index profiles. The outer boundary truncates fields; re-evaluate domain and mesh convergence for changed conditions.
