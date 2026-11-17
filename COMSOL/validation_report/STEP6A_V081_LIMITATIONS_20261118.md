# STEP 6a — 별도 한계 및 미지원 항목 보고 (v0.8.1)

일자: 2026-11-18 (KST).
**상태: 구현 및 스모크 테스트, 광학 손실 정량 검증 미완료.**

| 구분 | 현재 구현 한계 | 추가 증거/필요자료 |
| --- | --- | --- |
| 굽힘 모델 | 2D conformal equivalent-index weak-curvature 근사이며 3D 실변형이 아님. COMSOL 공개 step_index_fiber_bend의 2D axisymmetric bending 모델과 식·경계가 다름 | 동일 조건/좌표계에서 COMSOL 비교, 벤드 기하학 수치 정합성 |
| R_eff 응력광학 | 기본 stress correction 없음. 사용자가 f_literature와 출처 URL 및 조건을 제공해야 함 | 파장·유리 도핑·p11/p12·Poisson·온도/하중 조건 및 적용 모델 |
| PML / 누설 | 벤드로 유도된 방향성 방사와 원형 PML가 충분히 흡수하는지 미검증 | mesh/PML/domain radial/directional sensitivity + overlap-verified candidate |
| 물리적 손실 | 복소 neff의 Im는 radiation + material absorption + numerical error. 양의/음의 값 강제 보정 금지 | 흡수 포함 여부 분리, 비정상 부호 분석 및 정량 수렴 |
| dB/turn | 일정 R의 원형 전파만 가정. 굽힘 진입·진출/접속 scattering과 3D 변형 무시 | 실제 코일 시험 또는 3D propagation/coupling 모델 |
| 굽힘 방향 | 방향 θ의 등가 굴절률 변환만 사용; 유도 응력 광탄성/PM축 회전 미포함 | material tensor, stress tensor, 회전 및 polarization-verified modes |
| 모드 정체 | 공통 메시 complex Et overlap 사용하나 퇴화 모드/biorthogonal leaky modes 추적은 미지원 | 좌우 고유모드, subspace tracking, full-vector reference |
| V3 | 공식 COMSOL 문서에서 동일 비교용 모드별 정밀 Re/Im(neff) CSV 확보 못 함; 정량 오차표 미구현 | 사용자의 COMSOL exports 및 V3 동일 물리모델 비교 설정 |
| V4/V5 | PML, 공간 이산화·Im(neff) 수렴 기준 PASS 없음 | 각각 3+개 이상의 mesh/PML/도메인 변화 분석 및 허용오차 확인 |
| 웹 시각화 | 로컬 Python FEM JSON 결과만 표시. GitHub Pages에서는 벤드 FEM 실행하지 않음 | 브라우저 업로드/모바일/경계값 실제 UI 시험 |
| 다른 단계 | STEP6b~6f, STEP7 9개 validation suite 및 STEP8 계층 구조 전체 요청은 아직 미착수 | 사용자 별도 단계 승인 |

**반드시 표시:** computed_unverified / mode_tracking_failed / solve_failed / spurious_or_unstable_candidate. V3는 EVIDENCE_REQUIRED, V4/V5는 NOT_RUN으로 유지. 입력 해시가 안 맞을 때 과거 굽힘 곡선을 무효화한다.

SMOKE fixture는 단순 step-index와 작은 해석 영역으로 제작돼 정량 손실값/장거리 광섬유 제품설계에 사용해서는 안 된다.
