## v0.8.7 — 2026-11-19 — STEP5 vector Maxwell Bloch PWE

- Added independent transverse-H vector PWE with analytic circular inverse-epsilon Fourier coefficients and real nondispersive materials; retained scalar FDFD backend.
- Added versioned JSON Schema, 4 lattice motifs, fixed-beta bands, light lines, sampled-gap maps, actual cutoff/k sampling convergence data.
- Added gap-targeted finite triangular PBGF handoff with explicit computational-window policy and DOF guard. Real low-resolution finite FEM example is labelled unverified.
- Connected design/band/defect tabs, legacy export, shared design JSON and actual FEM/STEP4 result imports; stale-result invalidation and browser CI.

## v0.8.6 — 2026-11-19 — STEP4 physical integrals and field/subspace tracking

- Added full six-component material-labelled physical Gauss export; excluded PML by actual element label; raised quadratic-field quadrature to degree 8 for Et^4.
- Added precise quadrature Aeff/MFD2nd, sourced quartic n2/gamma, glass Poynting fraction, explicitly supplied radial PII and mode identities, signed loss/absorption distinction.
- Added adjacent field/subspace matching, parameter sweep tracking, D/S/ZDW with quartic local derivatives, dispersive ARROW root and conditional HOMER/surface candidates.
- Retained old v0.8.0 formulas/CLI; missing material/identity/convergence inputs remain TODO/null. No automatic HCF product certification.

## v0.8.5 — 2026-11-19 — STEP 3 mixed vector solver and applied tool integration

- Applied STEP1/2 web shortcuts; research panel remains accessible in HCF/vector view modes; app v0.8.5, scalar v0.7.1 unchanged.
- Retained mixed Maxwell assembly and companion backend; added N2/P2 and equivalent beta-squared sparse shift-invert backend with explicit resource budget.
- Fixed PML H reconstruction using transformed inverse permeability and requested SBC target; added optional target iteration with first-order boundary limitations.
- Added physical-domain 1 W normalization, finite six fields, QEP/pencil/dual-norm Gauss diagnostics, Stokes/mirror/longitudinal-content diagnostics.
- Kept independent degenerate modes; removed only field-rank duplicates; tracked fixed FE probe subspaces across remesh/window/PML changes.
- Added exact vector guided circular, absorbing PEC TE11 and outgoing-Hankel radiative references; reproducible V1/V4/V5 PASS_SCOPED benchmarks. Does not certify tiny-loss PCF/HCF or COMSOL parity.
- Added actual Python examples, mode selector and strict input/stale-result guards to web; full-index browser CI coverage.
- Preserved material provenance, Ge/F TODO blockers, original presets and existing scalar/Yee/Colab assets. Next PCF stage requires approval.

## v0.8.4 — 2026-11-19 — STEP 1·2 design/mesh verification

- Kept 25 preset geometric/optical numbers and design schema_version 1.0.0; added capability metadata and optional mesh policy controls.
- Added strict numeric/geometric preflight, nested JSON Schema rules, `inspect` and `mesh-study` CLI commands.
- Fixed y-major OCC ellipses by swapping semiaxes and rotating; no silent PCF hole clipping.
- Added actual-material wavelength sizing, curvature meshing, physical boundary labels, sampled circular-wall resolution and measured material/area/edge/aspect metrics.
- Compatibility policy reports bulk wavelength violations; explicit wavelength policy enforces the whole-domain target.
- Added actual mesh JSON preview, input-match indication and invalidation on edits; Python FEM remains external.
- Corrected Malitson source DOI only; coefficients and STEP 3 vector_fem.py unchanged.
- Validation: baseline 67 tests; final 83 tests; 25 default geometry generations, 75 sampled walls; no optical certification. Local browser launch blocked by runtime sockets; GitHub browser integration test added.

# Fiber Research Engine CHANGELOG

## v0.8.3 — 2026-11-18 — STEP 5 scalar Bloch prototype

- Independent 2D fixed-beta Bloch-Floquet scalar FDFD generalized Hermitian eigenproblem, including cross derivatives on 60° primitive cells.
- Triangular 1-hole, hexagonal honeycomb 2-site, square 1-hole and idealized Kagome 3-site circular-hole unit cells; strict non-overlap checks.
- Γ/M/K and Γ/X/M band structures, sampled reciprocal k-mesh gap candidates, normalized β/k0–Λ/λ map, air/host light line. No complete gap certification.
- Triangular PBGF 7/19-cell finite-section design JSON handoff, with existing vector FEM separately run. Other lattices blocked until CAD integration.
- Same GitHub Pages workbench 3 design tabs, input JSON export, Python results import, modal handoff export, source-mismatch invalidation. Real computed Python sample published; **default sample has 0 sampled gap candidates**, not fabricated.
- CI: all 67 Python tests PASS at initial integration and JS syntax check. Physical COMSOL V2/V4/V5/V6 and full-vector Floquet **NOT PASS**.
- Reports: STEP5_V083_DESIGN_VERIFICATION_20261118.md / STEP5_V083_LIMITATIONS_20261118.md.



## v0.8.2 — 2026-11-18 — STEP 6b ONLY

- generalized-plane-strain thermoelastic FEM with user-cited E,ν,α,ΔT; free axial resultant; no guessed Panda rod constants.
- isotropic strain-optic p11/p12 tensor (Q=ε_r^-1), spatial principal indices, core-area material B/beat-length proxy distinguished from modal B.
- Panda/Bow-tie stress_rod as separate physical domain; rod refractive index in presets explicitly a pure-silica surrogate, NOT commercial B2O3-doped rod data.
- anisotropic εxx/εxy/εyy/εzz triangle tensor mapped into existing HCurl/H1 vector FEM + radial PML; provisional x/y mode B only when clearly identified.
- pm_cli.py: design+material JSON -> Gmsh -> stress+photoelastic -> vector QEP -> source-hashed outputs; UTC timestamp, mesh SHA and engine v0.8.2.
- research_step6b_v082.mjs: source-controlled material editor, stress/principal-B map, material-vs-modal B separation, formula input and references, change invalidation.
- CI: 58 Python tests and JS syntax passed, including real small Panda Gmsh + stress + anisotropic vector eigen candidate. These are analytic/smoke tests, NOT high-precision PM benchmark.
- **Not completed:** COMSOL optical benchmark, realistic stress rod n(λ) and all composition-dependent elastic/p coefficients, mesh/PML convergence. STEP6c~6f + STEP7/8 entire work **not started** without approval.

[STEP6b implementation/verification](../../validation_report/STEP6B_V082_DESIGN_VERIFICATION_20261118.md) · [separate limits report](../../validation_report/STEP6B_V082_LIMITATIONS_20261118.md)

## v0.8.1 — 2026-11-18 — STEP 6a ONLY

- 굽힘 등가 굴절률: n_eq=n(x,y)[1+(x cosθ+y sinθ)/R_eff], x/y[µm], R[mm]. 사용자 문헌 URL/조건을 명시한 경우만 stress-optic R_eff 배율 선택.
- vector_fem.py 기존 복소 텐서/PML 결합에 bend epsilon correction 적용. 직선 계산 v0.8.0 및 디자인 스키마 1.0.0 보존.
- bending.py + cli.py bend: 반경/방향 공통 mesh 스윕, complex Et overlap identity, α_total[dB/m]/Δα/total dB per turn, 입력/settings SHA-256/UTC time, 분리된 solver error.
- 기존 웹에 research_step6a_v081.mjs 연결: 방향별 반경-손실 그래프와 계산근거·문헌 링크, 설계 JSON이 변경되면 결과 무효화.
- 테스트 42개 포함(기존 29 + STEP6a 12 + bend mixed-FEM 1) 및 JS syntax CI 성공. 단위/스모크 PASS는 물리적 V3/V4/V5 인증 아님.
- step-index PML smoke fixture 제공(실물 데이터/공인 COMSOL 수치 아님).
- V3 COMSOL axisymmetric 비교는 EVIDENCE_REQUIRED, V4/V5 NOT PASS. STEP6b~6f, STEP7 전체, STEP8 전체 미착수. 사용자 다음 승인 전 중단.

관련 자료: COMSOL/validation_report/STEP6A_V081_DESIGN_VERIFICATION_20261118.md 및 STEP6A_V081_LIMITATIONS_20261118.md.

## v0.8.0 — 2026-11-18 — STEP 4 ONLY

- physical_metrics.py 신규: 단위/부호 계약, signed α[dB/m], Et-only Aeff/second moment, radial Petermann-II, D/S/λ0, γ, HCF effective n₂와 glass fraction, HOMER, birefringence/beating, Δneff, ARROW resonance, anticrossing candidate.
- vector_fem.py: v0.8.0 engine_version, Et-only Aeff/MFD basis, α total dB/m+dB/km, optional full complex FEM quadrature export.
- step4_cli.py: 단일 모드 파일 후처리와 공통 메시를 사용한 N점 파장 full-vector FEM 스윕, 복소 Et-overlap Hungarian mode identity 확인.
- research_step4_v080.mjs: STEP4 JSON 시각화와 ARROW 이론선/FEM Re(neff) overlay, D dispersion chart. Browser FEM 실행 없음.
- unittest 및 CI: analytic Gaussian/polynomial/overlap/HOMER/ARROW/나노 범위 검증; 기존 프리셋·FEM smoke 보존.
- Aeff 정의 변경 주의: v0.8.0의 정확한 사용자 요구는 transverse Et (Ex/Ey)이며 과거 total-E Aeff와 구분해야 함.
- Design JSON 스키마 버전 1.0.0은 그대로, 계산 엔진 버전만 v0.8.0으로 기록.
- 미완료: STEP3 V1/V4/V5 물리 게이트, HCF confinement loss의 PML 독립 수렴, STEP5 Bloch bandgap 등. 승인 없이 다음 단계 실행 금지.

관련 보고서: COMSOL/validation_report/STEP4_V080_DESIGN_VERIFICATION_20261118.md 및 STEP4_V080_LIMITATIONS_20261118.md.

