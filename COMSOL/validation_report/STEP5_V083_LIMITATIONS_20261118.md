# STEP5 v0.8.3 — 별도 한계·TODO 보고서 (2026-11-18)

## 결론
STEP5는 **스칼라 Bloch 수치 엔진과 UI 설계 흐름**을 연결한 프로토타입이다. 물리적으로 인증된 PCF/PBGF/HCF 설계기가 아니며 STEP7 V2/V4/V5/V6 통과 전 '완료'라고 표시해서는 안 된다.

## 핵심 제약
1. 수치식은 scalar Ez surrogate. 실제 축방향 β가 있는 비균질 광섬유의 2D *vector Maxwell H(curl)* Bloch 고유해석이 아니다. TE/TM, 편광 분리, hybrid mode coupling 등 누락.
2. N×N structured finite difference로 원형 홀 경계가 계단형. full cell의 grid N 3점 이상 독립 수렴이 필요.
3. k-mesh는 reciprocal primitive parallelogram의 유한 격자이며 Γ/M/K 표본을 보탠다. 특정 path 및 유한 mesh에서 발견된 간극은 **SAMPLED_GAP_CANDIDATE**일 뿐, 전체 Brillouin zone 완전 gap이 아님. band count를 초과한 인접 밴드/수치 오차도 확인 필요.
4. n_air, n_host는 입력된 파장 무관 실수 상수. 실제 silica Sellmeier/도핑/흡수 및 계면 전계 연속 Maxwell 조건 미반영.
5. 'hexagonal'은 honeycomb 2-site lattice; 'triangular'은 1-site. Kagome는 이상화 원형 hole 3-site surrogate, 실제 Kagome strut/hypocycloid 형상이 아님.
6. 결함 설계는 triangular lattice에 한해 기존 finite PBGF 7/19 cell schema와 연동. Honeycomb/square/Kagome finite core geometry 및 면밀한 mesh 일치 미구현.
7. 기존 full-section vector FEM에서 모드가 발견되어도 해당 모드가 진짜 bandgap core confinement 또는 적절한 Im(neff) 수렴을 충족하는지 별도 V2/V4/V5 판정 필요.
8. 웹은 JSON 결과 표시기 + 설계 파일 생성기. 실제 SciPy/Gmsh full solver를 GitHub Pages 브라우저에서 실행하지 않음.
9. α confinement, MFD, full vector core modal loss, 모드 band crossing follow 등은 STEP5에서 정량 비교 완료되지 않음.

## 후속 검증 게이트
- β/ω/k 변환 및 Bloch 위상 부호를 이론 plane-wave 추가 케이스로 교차 점검.
- fixed β에서 full-vector PCF Floquet eigenproblem 구현 후 COMSOL Official Band-Gap Analysis 설정 동치 비교.
- grid N, k-mesh, number of bands, dielectric boundary refinement 각각 독립 스윕 + err tolerance 사전 공지.
- Sampled gap이 true gap인지 dispersion-following, dense BZ sampling, plane wave expansion 독립 솔버로 검증.
- 7/19 core defect field confinement/positive loss/PML independent 등 V2/V4/V5 게이트 연결.
- HCF Kagome 실제 물성·두께·strut geometry와 면 모드 PBG 검증은 별개 연구.

인증되지 않은 문헌/COMSOL 수치를 인위적으로 채우지 않는다. 사용자 요청 STEP6 c–f, STEP7 종합 탭, STEP8 상세 UI 구조는 이 STEP5 작업에서 새로 구현하지 않는다.
