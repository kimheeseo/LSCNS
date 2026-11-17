# STEP 4 v0.8.0 — 별도 한계 및 필요 근거 보고서

일자: 2026-11-18 KST. 분류 기준: 구현 여부 / 공식 단위 테스트 / FEM 물리 정확도 / 상용제품 인증은 서로 다름.

| 물리량 | 제한 사항 | 검증에 필요한 자료 |
| --- | --- | --- |
| HCF α total vs confinement | FEM Im(neff)는 재료흡수와 누설을 합산; STEP3 PML 및 모드 수렴 미통과 | 문헌 동일 HCF 구조, h_wall 충분한 메시, 3수준 mesh, PML 두께/강도/차수, 반경/도메인 sweep |
| Petermann-II | radial scalar 근접계는 구현, 일반 2D full-vector gradient는 PII-like만 | IEC 60793-1-45 near/far field 정의, 실제 필드·적분 창 수렴 |
| D,S, λ0 | 5점 cubic fit 및 overlap identity 검증만; 실제 모드 branch/퇴화 가능 | 실측 또는 공식 Sellmeier n(λ), 2개 이상의 Δλ, full vector field overlap / subspace tracking |
| γ | n₂ 미지이면 TODO, Et scalar 근사 | 공인 n₂(λ), 재료 조성·온도·파장 |
| HCF n₂ effective | air/gas/glass n₂ 둘 다 필요. 파워비율만으로 nonlinear n₂를 대체할 수 없음 | 가스 종류, 압력·온도, 밀도/광파장, n₂ 표 및 영역별 field |
| 유리 power fraction | 계산 모듈은 존재하나 FEM 결과의 physical-domain glass mask 기본 연결 없음 | gmsh physical group ↔ FEM triangle quadrature 매핑 및 정상 Poynting 테스트 |
| HOMER | 양의 수렴 α_FM, α_HOM와 모드 ID 검증 없으면 TODO | 개별 HOM/Fundamental 필드 지도, eigenmode overlap, PML 손실 검증 |
| anticrossing | 두 mode의 gap minimum + glass fraction 교차는 후보 판별에 불과 | glass surface-localized modes field pattern, 세밀한 λ mesh, 물리 branch 교차 검사 |
| ARROW | t,n_glass,n_air 고정 조건의 1D Fabry–Perot 근사 | SEM으로 확인한 벽두께, n_glass(λ), n_air(λ), 실제 FEM loss spectrum |
| B/LB | 단순 x/y intensity fraction은 PM 응력 복굴절 보증 불가 | Panda/bow-tie stress tensor, photoelastic p_ij, thermoelastic boundary, 모드쌍 검증 |
| 모드 overlap | 동일 메시 및 centroid quadrature 사용, leaky 비정규 모드 biorthogonal overlap 미구현 | remesh interpolation, left/right eigenfield, 퇴화모드 subspace principal angle 검증 |
| 웹 | 정적 JSON 결과 뷰. FEM 재실행·실제 pixel screenshot QA 미완 | 배포 완료 후 브라우저 파일 업로드/오버레이 수동 회귀 |

### 잔여 위험
- 기존 STEP3 V1/V4/V5는 사용자가 제시한 독립 검증 기준 PASS가 확인되지 않았으며 이를 건너뛰고 수치 인증을 주장하지 않는다.
- synthetic analytic unit tests 성공은 물리/제품 정밀도 승인으로 해석하지 않는다.
- 노출된 Aeff는 새 Et 기준이므로 이전 총 E(Ex,Ey,Ez) Aeff와 비교하려면 동일 정의로 다시 계산한다.
- air n=1.00027은 기존 설계의 단순 상수 가정이지 P/T/분산의 실측 모델이 아니다.
- PBGF Bloch bandgap, Floquet unit-cell, Kagome exact CAD, surface optical stress tensor, 6a..6e는 이번 STEP4 범위 밖이다.

### 다음 제안 (승인 전 코드 작업 금지)
STEP 3 V1/V4/V5의 근거 자료와 명확한 정량 합격 기준을 확보하고, 정량 수렴 시험/별도 한계 보고까지 마친 뒤 사용자에게 보고한다. 그다음 사용자가 별도로 승인한 순서에 한해서 PCF V2/V6, STEP5로 진행한다.
