# Fiber Research Engine CHANGELOG

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
