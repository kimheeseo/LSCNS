# STEP 6b v0.8.2 — 별도 한계 / 증빙 필요 자료

**상태:** 열응력·광탄성 공식/단위 테스트 및 소형 FEM smoke 확인. 상용 PM fibre 제품 물리 검증 미완료.

| 한계 | 설명 | 필요한 자료 |
| --- | --- | --- |
| 응력봉 굴절률 | 현 Panda/Bowtie rod n0는 순수 SiO2 **광학 대리값**, 실제 stress-rod 도핑의 n(λ)은 알려지지 않음 | 제품 원형 SEM·응력봉 조성/B2O3/Ge/F Sellmeier 및 실제 n(λ) |
| 기계 상수 | E,ν,α(T), 냉각 ΔT 및 stress-free T 제품별 상이 | rod/core/clad 조성별 문헌/실측값과 공식 출처, T 이력 |
| 잔류 응력 | generalized plane strain 균일 εzz와 force-free, 외주 traction free. 점탄성/어닐/코팅·장축 warping/3D 고려 안함 | COMSOL GPS 모델과 열처리·점탄성·코팅 데이터 |
| 광탄성 자료 | p11,p12 strain-optic 무차원 사용; stress-optic B1/B2(m²/N)와 혼용 금지. intrinsic dn/dT 미포함 | 파장·온도 의존 p_ij, intrinsic thermo-optic coeff |
| geometry | 현재 bow-tie의 rod는 타원, 실제 arc/wedge CAD 아님 | 실제 도면/SEM, mesh conformity 및 geometry convergence |
| 물리 모드 복굴절 | core material Q 평균 ≠ full-vector modal B. x/y candidate 분류만으론 HE11 근본모드 정체 검증 불충분 | x/y FEA mode overlap+polarization, 독립 COMSOL neff·B, 측정 LB |
| 벡터 tensor PML | diagonal/offdiagonal eps를 복소 PML에 더함. spurious eigen, numerical PML mesh sensitivity 및 leaky mode left/right biorthogonal tracking 미검증 | independent formulation, 3+ mesh/PML curves, residual/orthogonality |
| 전계 수치 | n_eff 차이 매우 작으면 현재 Nédélec-1 mixed QEP의 numerical noise와 동급 가능 | h/p refinement, reference high-order FEM, B accuracy > numerical floor |
| 입력 provenance | URL 필수만 기계적 확인. URL 문서의 실제 계수·조성까지 자동 검증한 것은 아님 | 출처 원문, 표/page, 적용 조건 수동 검증 |
| 웹 | Browser JSON 편집·지도 시각화, 실제 Python FEM은 로컬/Colab만. 모바일/브라우저 시각 검증 별도 | Pages 배포 및 E2E 접속 검사 |
| 기타 | STEP6c–6f, STEP7 V1–V9 전체, STEP8 전체 리팩토링 미착수 | 다음 단계 개별 승인 |

**검증 게이트:** 정확한 stress-rod 물성/형상 확보 → 독립 COMSOL generalized plane strain σ/ε 검증 → 광탄성 εr 비교 → 동일 geometry/mesh/PML neff_x/neff_y 확인 → 실측 B/LB 오차 보고. 기준값과 허용오차는 사용자/공식 검증 요구가 없는 한 임의 설정·PASS 처리 불가.

**다음 제안 (진행하지 않음):** STEP6c. 신뢰 가능한 구조/광학 상수를 공급받아 supermode와 결합계수 κ를 계산하고, 실험/공식 기준 데이터와 비교. STEP6d/e/f나 STEP7/8은 승인 없이 구현하지 않는다.
