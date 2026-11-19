# STEP7 v0.8.9 — 검증 한계 / 다음 제안

| 항목 | 통과를 막는 원인 | 필요한 자료·후속 계산 |
|---|---|---|
| V2 | 공식 10모드 neff 표는 확인했지만 원 parameters.txt의 형상·재료·파장·PML 전체 입력 미확인 | COMSOL Application Library 원본 `microstructured_optical_fiber_parameters.txt` 및 가능하면 동일 버전 field/complex neff export. 그림에서 추정하지 않음 |
| V3 | 현재 conformal 근사와 COMSOL 2D axisymmetric eigenvalue는 정식화·계산창·r0 normalization이 다름. 공개 본문의 rounded straight neff는 bent complex reference가 아님 | COMSOL full precision bent complex eigenvalue, r0/βr0 normalization, mesh/PML·geometry, equivalent-formulation 비교 정책 |
| V6 | 실제 FSM의 cutoff 2/3/4 변화가 남고 일부 λ에서 localized FM 식별 실패. HOM cutoff 및 quantitative reference curve 비교 없음 | higher-cutoff vector PWE, more FM/HOM target shifts·overlap identity, ring/mesh/PML/λ 수렴, primary reference digitization와 불확도. finite V<π만으로 ESM 인증 금지 |
| V7 | fixed common probes에서 2편광 core-like subspace와 동일 branch·내부 loss peak를 전체 파장에 걸쳐 안정적으로 확보하지 못함 | adapted core-mode search, mesh/PML·창 변화, resonance spectrum refinement, actual field identity. FAIL은 현재 benchmark 검증 미통과이며 ARROW 이론 자체의 반증 아님 |
| V9 | CSV 수치 일치는 물리 입력·모드·단위 정규화 일치의 증명이 아님 | source design/COMSOL model manifest, fields/mode labels, explicit time convention and tolerances; actual exported CSV needed for external evidence |

V1은 explicit circular HE11의 선택 V 구간, V4/V5는 strongly leaky circular synthetic reference에 한정한다. 해당 PASS를 HCF 극소손실이나 arbitrary PCF/PM/MCF 품질 인증으로 확대하지 않는다. V8/V9 PASS는 독립 공식/파서 기능 검사다.

함께 요청하신 STEP8은 STEP7 결과를 먼저 보고한 뒤 별도 소프트웨어 작업으로 진행했다. `lswave/` 아래 안정 facade와 결과 계약을 추가하며 물리 검증 등급은 변경하지 않는다. 다음 물리 작업은 V2/V3 원본 자료 확보 → V6/V7 모드 식별·수렴 개선 → STEP7 gate 재평가 순서다. [STEP8 구조·UI 보고서](STEP8_V0810_VERIFICATION_20261120.md)를 참고한다.
