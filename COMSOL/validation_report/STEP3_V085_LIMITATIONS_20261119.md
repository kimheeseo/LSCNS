# STEP 3 v0.8.5 — 한계·미확인 사항

- V1/V4/V5는 명시한 원형 합성 조건에서 PASS_SCOPED다. 임의 PCF, PBGF, NANF/DNANF의 극소 Im(neff), 실제 손실 dB/km, COMSOL 동등성은 **미인증**이다. 고손실 outgoing Hankel 시험을 저손실 HCF 수치 정확도로 외삽하지 않는다.
- 실제 COMSOL 모델 또는 동등한 독립 연구 문헌의 전체 형상/경계/재료/meshing 수치와 비교하는 PCF V2/V6, HCF V7은 다음 승인 단계다. STEP 0 감사표의 미확인 항목은 임의로 채우지 않았다.
- Ge/F 도핑 Sellmeier 검증 계수는 TODO로 차단한다. 기존 공기 constant는 환경 조건 미지 가정이다. 자유 입력 굴절률이 물리적으로 타당하다는 보장은 없다.
- N1/P1 기본 정확도는 N2/P2 benchmark와 다르다. 예시의 PASS를 모든 element_order, preset 또는 mesh policy에 적용할 수 없다. 기존 25개 boundary_refined preset의 벌크 파장 기준 미달 보고도 유지된다.
- single-machine SciPy sparse LU의 메모리/시간이 제한적이다. max_dofs 상향은 성공/안전한 메모리 사용을 보장하지 않는다. 3D/분산 MPI/SLEPc는 TODO이다. 브라우저는 벡터 계산을 하지 않는다.
- PML은 원형 전용. 1차 SBC는 COMSOL 원통형 고차 SBC나 DtN과 동일하지 않다. SBC β 반복 자기일치성은 loss 인증이 아니다. 복소 고유계의 모든 가짜 모드를 자동 배제한다는 수학적 보증은 없다.
- 자동 분류는 편광/거울 대칭/종방향 장 함량 진단이다. 불연속 경계 근처 희소 probe, 우연한 모드 혼합, 고차/다중 분기 추적 한계가 있다. 일반 study의 추적 부분공간은 첫 두 모드이며 C6v 모드 분류/대칭 축약은 TODO이다.
- 웹의 필드 산점도는 요소 평균의 희소 표시다. 실제 전체 복소 FEM 계수를 웹에서 복원하지 않는다. MFD/Aeff의 기본 cell-average 진단과 정밀 full quadrature 후처리를 구분해야 한다. γ/도핑 비선형 계수·Petermann II 등을 새로 인증하지 않았다.
- 기존 Kagome/conjoined/bow-tie 등의 명칭과 근사 CAD 차이, 응력광학 일반 xz/yz tensor 미지원, 기존 PM/굽힘/스칼라 Bloch의 검증 한계를 유지한다. 이번 작업에서 후속 단계의 구현 범위를 늘리지 않았다.
- 결과 JSON의 benchmark 메타데이터는 데이터 파일의 출처 추적용이며 전자 서명이나 임의 수정에 대한 인증이 아니다. 입력이 다르면 웹은 표시를 거부하고, 입력 편집 시 결과를 무효화한다.

다음 제안은 **고체 코어 PCF 검증(V2/V6)**이다. 승인 전에는 진행하지 않는다.
