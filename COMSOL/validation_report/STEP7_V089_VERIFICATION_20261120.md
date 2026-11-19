# STEP7 v0.8.9 검증표 — 2026-11-20

**전체 물리 검증은 미완료다.** `all_required_passed=false`, STEP8 구조·UI 개선은 명시된 별도 요청으로 진행하며 이 물리 판정은 승격시키지 않는다. software tests/CI 성공을 V1–V9 전체 물리 PASS로 표시하지 않는다.

| 항목 | 판정 | 실제 결과 / 적용 범위 | 허용오차·필요 근거 |
|---|---|---|---|
| V1 | PASS | exact HE11 b–V 4점(V=1.8,2.0,2.2,2.35); 최대 neff 오차 2.53573321 ppm | ≤10ppm; explicit nc=1.5/ncl=1.44 synthetic HE11 interval only |
| V2 | 근거 필요 | 공식 Table1의 10개 PML neff를 확인, Im 시간 관례 변환 명시 | 원본 microstructured_optical_fiber_parameters.txt 전체 입력 미확인. 기준 오차값 미정 |
| V3 | 근거 필요 | 공식 core/clad/λ/R/크기 확인; full precision bent complex neff 미확인 | axisymmetric neff×r0 정규화·정식화·창이 conformal 근사와 다름 |
| V4 | PASS | fresh 실제 outgoing circular reference, 두께/강도/차수 각각 3값, 동일 FE subspace | 각 sweep Im 변동 ≤5%, min positive Im>1e−10. tiny-loss HCF 인증 아님 |
| V5 | PASS | guided/radiative mesh 각 3수준, actual edge_max/min 및 요소 수 확인 | 마지막 centroid ΔRe≤1e−5, radiative relative ΔIm≤5% |
| V6 | 근거 필요 | actual Γ-cell vector Maxwell fixed-λ FSM β root, PWE cutoff 2/3/4, finite PCF FEM 후보 | FSM 잔여 cutoff 변화·일부 λ FM 식별 실패·HOM cutoff/문헌 정량 비교 미검증. ESM 증명 아님 |
| V7 | FAIL | actual single-ring ARF 7λ 스캔, core-like 2편광 부분공간 추적 실패 구간 존재 | 신뢰할 수 있는 내부 loss peak 미확보. ARROW 식 자체를 계산해 PASS로 대체하지 않음 |
| V8 | PASS | independent SI dB/m↔dB/km, 음수 Im 보존, D 부호, Gaussian Aeff | loss 1e−12 relative; D 1e−5 ps/nm/km abs; Aeff 1e−10 relative |
| V9 | PASS (기능 시험) | COMSOL 주석·복소수/분리 열·D지수·단위·부호·행 매칭·오차표; actual vector result 연결 | 합성 parser/비교 시험 PASS. 실제 COMSOL CSV의 물리 동일성은 별도 검증 |

공학적 10ppm/5% 등은 선언된 수용 정책이며 논문/COMSOL의 보증 오차를 임의로 인용한 값이 아니다. V2/V3/V6에서 미확인 허용값과 자료는 null/근거 필요로 남긴다.

로컬 Python **134 tests PASS**, Node **13 tests PASS**. 신규 suite의 actual 계산과 별도로 기존 scalar/vector/STEP1–6 회귀를 실행했다. 실제 브라우저 및 GitHub CI 결과는 후속 증거로 기록한다.

[원 결과](step7_v089/validation.json) · [설계/판정 정책](STEP7_V089_DESIGN_20261120.md) · [실행·CSV·출처](../universal_silica_v01/fiber_research_v1/STEP7_V089.md) · [한계/후속 계획](STEP7_V089_LIMITATIONS_20261120.md) · [Python 로그](step7_v089/python_tests.txt).

V4/V5 원 데이터는 fresh STEP3 실제 재계산에서 가져왔으며 이전 보고서 번호를 통과 배지로 단순 변경하지 않았다. V5 mesh input hash와 실제 재생성 요소 수를 대조했다. V6는 n_glass를 n_FSM으로 치환하지 않으며 V7의 capillary 값은 shift seed일 뿐 기준답이 아니다.
