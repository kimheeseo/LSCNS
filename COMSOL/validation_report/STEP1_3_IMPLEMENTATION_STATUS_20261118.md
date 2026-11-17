# STEP 1–3 개발 반영 및 검증 게이트 — 2026-11-18

STEP 0 원문: [LS_Wave_Optics_STEP0_Audit_20261118.md](./LS_Wave_Optics_STEP0_Audit_20261118.md).

## 반영 경로
- Python 독립 모듈: [fiber_research_v1](../universal_silica_v01/fiber_research_v1/) — 25종 프리셋 JSON, 버전형 schema, 저장/불러오기, gmsh OCC geometry/physical domains/PML annulus, Nédélec-H1 quadratic-beta experimental FEM, CLI, 자동 민감도 실험 스크립트.
- 웹 결과/프리셋: [research_fiber_v1.mjs](../universal_silica_v01/web/research_fiber_v1.mjs) — 기존 웹 index.html에 독립 패널 추가. 브라우저는 Python FEM을 실행하지 않고 결과 JSON만 표시함.
- CI: [.github/workflows/comsol_fiber_research_v1.yml](../../../.github/workflows/comsol_fiber_research_v1.yml) — Python 3.11, gmsh/scikit-fem 설치 후 schema와 meshing/solver smoke 9개 성공(다만 물리 검증 아님).

## 단계별 상태
| 단계 | 실제 코드 | 미완료/미검증 |
| --- | --- | --- |
| STEP 1 | A–H / 25 JSON 디자인, 버전/범위 검사, 저장/읽기, 웹 JSON editor | FEM 재료 모형의 정확한 Ge/F 도핑 계수, 모든 모델의 상용사양 보증 불가 |
| STEP 2 | gmsh OCC 단면 분할, 물리 도메인, 원형 PML shell, 국소 메시 기준, 품질 통계/보고서 | CAD 상세형상(Kagome hypocycloid/conjoined 등) 일부 근사, C6v 조건 미구현, 물리 메시 수렴 미검증 |
| STEP 3 | Python HCurl(Nédélec1)+H1(P1) 결합, sparse QEP shift-invert, complex tensor radial PML, 1차 SBC 근사, 6성분 샘플, 약형 Gauss/고유잔차, parametric study | 작은 Step-index 메시/벡터 고유값 경로 1건 CI 실제 실행. 외부 HE11/COMSOL 비교, spurious-mode 배제, Im(neff) 정량 및 PML convergence 여전히 미검증 |

## 절대 유의
기존 브라우저 Yee 실험의 neff 수렴 실패 및 HCF loss 미검증은 이 새 코드가 존재한다고 자동으로 해결되지 않는다. 제품 스펙 인증이나 COMSOL 등가 구현으로 간주하지 않는다.

## GitHub Actions 독립 실행 기록
- [COMSOL fiber research v1 — 성공한 CI 실행 #44133166704](https://github.com/kimheeseo/LSCNS/actions/runs/44133166704) (2026-11-18 KST). Python 3.11, gmsh 4.15.2, scikit-fem 12.0.2, SciPy 1.17.1.
- `python -m compileall`, `unittest` 총 **9/9 성공**: A–H의 25프리셋 범위 검사 및 JSON 왕복, 작은 원형 Step-index gmsh 메시 생성, mixed vector 고유모드 후보 1건, E/H 6성분·잔차 JSON 구조, 복소 QEP 결합의 허수부 보존.
- 최초 실행에서 scikit-fem의 `BilinearForm` 기본 실수 배열이 복소 계수를 버리는 경고를 확인하였고 **`dtype=np.complex128`을 지정하여 수정**. 수정 후 CI는 성공했다.
- 이것은 **스모크 테스트 성공**이며 공인 광학 정합성 시험 통과가 아니다. HE11 독립 수치 오차, 모드 중복/스퓨리어스 제거, 메쉬/PML 손실 수렴, 모든 A–H CAD 특성 검증, C6v/SBC 엄밀도는 아직 PASS 아님.
- 웹 기존 Yee와 Python 벡터 FEM은 알고리즘이 다르고 두 방식 모두 자체의 검증 제한이 있다. 산출 `mode_result.json`의 `loss_dB_km_certified=false`를 반드시 유지한다.

## 미구현 항목의 정확한 의미
- Endlessly single-mode: **근사 문헌 조건을 참고한 도식만 제공**, 고차 모드 컷오프 해석으로 인증하지 않음.
- PBGF 밴드갭: 유한 단면 지오메트리만 존재하며 **ω(k) Floquet 고유주파수 맵 미구현**.
- Kagome/negative-curvature/conjoined: 실제 hypocycloid/복합 호 CAD 대신 단순 튜브로 표시한 부분 존재.
- PM: Panda/bow-tie 형상만 존재, 기계적 응력·광탄성 이방성 해석 미구현.
- C6v 대칭 및 정밀한 Scattering BC: 옵션 검증이 완료되지 않아 C6v는 예외, SBC는 Robin 근사.
- Ge/F 도핑 Sellmeier: 조성별 공식 계수 출처 미확정으로 예외를 발생시키며 추정 계수는 삽입하지 않음.

**GitHub에 코드가 반영되었음과 전 계열 Full-vector 물리 모델 검증 완료는 서로 다릅니다.**
