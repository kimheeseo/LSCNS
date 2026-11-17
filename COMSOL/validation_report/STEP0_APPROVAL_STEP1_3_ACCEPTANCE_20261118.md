# STEP 0 승인 및 STEP 1–3 구현·검증 기준 — 2026-11-18

원문 감사: [LS_Wave_Optics_STEP0_Audit_20261118.md](./LS_Wave_Optics_STEP0_Audit_20261118.md).

## 승인 범위와 보존 원칙
- STEP 0의 읽기 전용 감사 결과를 승인된 기준선으로 보존한다. 기존 수치 기록을 새로 실행한 결과로 표시하지 않는다.
- 기존 `COMSOL/universal_silica_v01/web/`의 scalar radial P1 FEM, Yee-FDFD, Colab 결과 가져오기, HCF/MCF·CAD·프로젝트 UI 및 기존 schema 3 프로젝트를 변경하거나 무효화하지 않는다.
- 신규 구현 위치: `COMSOL/universal_silica_v01/fiber_research_v1/` (Python+JSON) 및 선택적 웹 결과 뷰어. Python 실행과 정적 GitHub Pages를 분리한다.
- 본 문서는 **승인/검증 계획**이며 수치해석 완료 또는 COMSOL 결과 일치를 의미하지 않는다.

## STEP 1 — 스키마 및 프리셋
8개 계열: index-guided(step/graded/trench/W/multilayer), solid-core PCF(hex/square/LMA/highly-nonlinear/elliptic holes), PBGF(7/19-cell), Kagome, AR-HCF(single/NANF/DNANF/conjoined/negative curvature), Bragg, MCF(hex/ring/trench), PM(Panda/bow-tie/elliptic core). 프리셋은 **기하학 초기값**으로만 표시하며 보증된 제품 또는 물리 계산 결과로 오인하지 않는다.

- 필수: `schema_version`, `family`, `variant`, `wavelength_nm`, `parameters`, `materials`, `mesh`, `boundary`.
- JSON 저장/불러오기 및 오염·범위·미지원 파라미터에 대한 명시적 예외, 기존 UI 프로젝트와의 분리.
- `endlessly single-mode` 도식·PBGF bandgap은 **밴드구조 별도 ω(k) Bloch 계산 전에는 단정 금지**. 단면 미리보기를 밴드갭 맵으로 표시하지 않는다.
- Ge/F 도핑 Sellmeier 계수는 신뢰 가능한 특정 조성 출처 없으면 TODO; 임의 계수 보간 금지. 응력 Panda/보타이 단면과 광탄성 복굴절 예측을 구분한다.

## STEP 2 — gmsh
- 재료별 2D physical domain, 바깥 원형 PML 영역/태그, 외곽 PEC 또는 선택적 SBC 설정, 메시 형상·경계 표시.
- ARF 벽에서 `h_wall <= thickness_um / through_wall_elements`; 곡률·파장당 요소 수도 하한 기준에 포함. 국소적 최소 요소를 사용해 전체 도메인을 지나치게 고밀도로 만들지 않음.
- 2D 삼각형 최소/최대 에지, 종횡비, 요소 수, 면적, 메모리 추정 기록; 대칭축이 유효한 경우에만 C6v 옵션 허용, 기본은 전체 도메인.
- grid 및 PML 매개변수 수렴은 서로 독립 검증. 부호가 바뀌거나 0 근처 Im(neff)가 불안정하면 손실 **미검증** 처리.

## STEP 3 — 실제 Python 벡터 계산 합격 기준
- 혼합 `HCurl(E_t)+H1(E_z)` 또는 공간 이산화·제약 처리 근거가 명시된 동등 대안. 스칼라 FEM과 구분. 목표 neff의 sparse shift-invert, PEC/PML/SBC 구현 근거와 사용 버전 명시.
- 복소 neff, Ex/Ey/Ez/Hx/Hy/Hz, 모드 편광·대칭, 상대 residual, 유한성, 경계/PML 모드 판별, 메시/PML sensitivity 결과 JSON.
- `exp(iβz−iωt)`에서 양의 Im(neff)가 수동 손실. `α[dB/km]=(20/ln10)(2π/λ[m]) Im(neff) × 1000`. 부호 절댓값으로 보정 금지. 흡수·방사 분리 표기.
- 기준: 원형 step-index HE11 독립 특성 방정식, scalar 약유도 비교(서로 다른 oracle), 해상도 3단계, PML 두께/강도/차수 3방향, 도메인 크기, 양의 손실, 물리영역 모드 중첩.
- 구현 중인 기능과 **실행/검증 통과**를 구분한다. gmsh/scikit-fem 의존성 설치 및 실제 재현 시험 전에는 모델의 수치 결과를 사용 금지.

## 검증 상태 (2026-11-18 등록 시점)
- 기존 2026-11-12 웹 Yee 결과: neff grid 수렴 **미통과** (N72→96 차이 5.549e−6 > 1e−6).
- 기존 PML: 직선 누설손실 −7.743~14.963 dB/km, 부호 변화로 **미통과**.
- 기존 HCF 0.45 µm wall, grid 1.5~2.25 µm: wall 해상도 부족으로 **미검증**.
- 새 STEP 1–3 코드는 기능별 실행 증거/수치 결과가 생길 때 별도 보고서에만 PASS를 추가한다.

공식 참고: [COMSOL Microstructured Optical Fiber](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html), [COMSOL Photonic Crystal Bandgap](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.bandgap_photonic_crystal/bandgap_photonic_crystal.html).
