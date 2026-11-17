# Fiber Research Engine CHANGELOG

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
