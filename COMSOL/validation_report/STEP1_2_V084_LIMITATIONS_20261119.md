# STEP 1·2 v0.8.4 — 한계와 다음 단계 (2026-11-19)

## 현재 범위

입력·기하·메시의 수치 검사를 추가했다. 기존 스칼라 웹, Yee/FDFD, Colab import, 벡터 고유계·PM·굽힘의 물리 코드는 보존했다. JSON 프리셋의 숫자는 유지하며 메타데이터만 추가했다. 검증 83개 통과를 새로운 광학 물리 인증으로 해석하지 않는다.

## 미완료/미검증

1. **정확한 특수 CAD:** 기존 Kagome/hypocycloid, conjoined/negative-curvature, bow-tie, PBGF 원형 코어/surround, 동심 nested tube는 명칭의 모든 실제 섬유 구조를 나타내지 않는다. graded/multilayer는 기존 교대 상수 링이며 임의 층별 물성과 연속 프로파일은 별도다. capabilities와 inspect 출력에 공개한다. 정확한 구조 도면·정의가 필요하다.
2. **Endlessly single-mode:** 기존 웹 0.406 도표는 단일 결함·원형 공기공·triangular silica PCF의 문헌 참고 구역이다. 유한 링 cutoff/고차모드 억압 인증이 아니며 square·타원·다중 결함에는 적용하지 않는다. 정량 출처는 Folkenberg et al. (2003), https://arxiv.org/abs/physics/0307009 로 연결했다. 새로운 모드 계산은 수행하지 않았다.
3. **기존 메시 정책:** boundary_refined는 기존 벌크 크기를 보존한다. 실제 생성한 기본 25개 모두 벌크의 엄격 파장 기준 FAIL_POLICY가 보고된다. wavelength는 명시적으로 선택해야 하고 계산량이 커진다. 새 규칙이 기존 solver의 22,000 mixed DOF 상한을 해제하지 않는다.
4. **메시 검사는 국소 표본·정책이다:** 법선 교차는 원형 annulus에만 적용하고 기본 24개 각도를 표본 검사한다. 임의 벽 전체 보증은 아니다. 재료별 phase λ/n 기준은 PML 복소 stretching·흡수 skin depth·장 기반 적응 refinement를 인증하지 않는다. 0.8 sizing factor, 1.35 maximum-edge allowance, 0.99 coverage, 2% area threshold는 공개된 수치 검사 정책이다.
5. **C6v:** 전체 도메인이 기본이며 대칭 축소는 기존처럼 명시적 미지원 오류를 반환한다. 검증되지 않은 경계조건을 추가하지 않았다.
6. **물성:** 순수 실리카 Malitson DOI 메타데이터만 교정했다(계수 그대로). air n=1.00027, constant 물성은 기존 가정이다. 실제 Ge/F 도핑·PM rod·n2·가스 조건은 출처가 있어야 한다.
7. **수렴:** mesh-study는 3개 실제 메시의 요소·벽·CAD 면적 변화를 출력한다. 면적 오차의 비단조 변화를 그대로 남긴다. neff·Im(neff)·PML 손실 수렴을 계산하지 않았다.
8. **웹 미리보기:** 최대 6,000개 삼각형의 결정적 표본과 모든 라벨 경계를 보여준다. 정확한 전체 메시 파일은 mesh.npz/geometry.msh다. 브라우저 FEM은 추가하지 않았다. 로컬 Chromium은 socket 실행 제한 때문에 화면 검증을 완료하지 못했다. GitHub CI의 별도 브라우저 검사를 확인해야 한다.

## 다음 제안 — STEP 3 보완·검증 (승인 전 미착수)

기존 HCurl/H1 코어를 재작성하지 않고 HE11 독립 기준해, 퇴화 편광 중복 제거, Gauss·고유잔차·PML 모드 선별, field-overlap 모드 추적과 메시/창/PML 수렴부터 검증한다. 현재 eigenvalue-only dedup, fixed-target SBC, 일반 tensor 범위는 이 단계에서 다룬다. V번호 정의·비교량·허용오차를 먼저 문서화해야 한다.

정확한 CAD의 미완료 항목과 C6v는 기능 목록에서 계속 TODO로 유지하며, 이번 보완만으로 A–H 모든 변형의 기하·광학 동등성이나 COMSOL 대체 완료를 선언하지 않는다.
