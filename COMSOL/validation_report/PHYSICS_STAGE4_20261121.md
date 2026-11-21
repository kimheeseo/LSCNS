# 4단계 — 전체 물리 gate 재평가 (2026-11-21)

구현·로컬 실행 완료. **all_required_passed=false**, strict physical gate의 종료코드는 **1**이다.

| 항목 | 재평가 결과 | 근거 범위 |
|---|---|---|
| V1/V4/V5 | PASS 유지 | 과거 합성 원형 HE11/PML/메시 검증 범위. 이번 단계에서 새로 전체 재계산한 것으로 표시하지 않음 |
| V2 | NEEDS_EVIDENCE | exact parameters 및 입력 일치·모드 장·독립 수렴·비교 정책 미확보 |
| V3 | NEEDS_EVIDENCE | 정규화 코드·공식 관계 확보. 실제 정밀 bent CSV와 두 반경·정식화 대응 미확보 |
| V6 | NEEDS_EVIDENCE | 실제 FSM cutoff 확대, FM 회복·메시 스윕·HOM 후보 보존. 수렴/식별/cutoff/정량 문헌 검증 미완료 |
| V7 | FAIL | 7점·중간점의 continuous branch 및 내부 loss peak 미확보 |
| V8/V9 | PASS 유지 | 독립 공식·합성 CSV 파서 기능 범위 |

## 새 자동 평가 경로

`physics_gate.py`와 `physics_20261121/external_manifest.json`을 추가했다. 원본 입력·CSV·모드 대응·수렴 증거·허용값을 파일 hash와 연결하며 결측을 0 오차나 PASS로 바꾸지 않는다.

- V2/V3: 원본 artifact SHA256, 명시적 부호, 실제 field/subspace 대응, 3수준 독립 수렴, CSV 정량 오차를 검사한다. V3는 raw bent 값과 실제 rAverage/r0로 reference를 다시 정규화하고 정식화 검토를 요구한다.
- V6: 실제 FM/FSM으로 VPCF 및 문헌 정량 오차를 계산한다. 사람이 기록한 0 오차나 residual만으로 수렴·HOM cutoff를 인정하지 않는다. 문헌 입력·불확도와 HOM 검토가 빠지면 보류한다.
- V7: gap/잘못된 spectrum/끝점 peak를 거절하고, 유효한 peak도 실제 장 식별·독립 수렴·3수준 spectral refinement 없이 PASS시키지 않는다. refinement 원 스펙트럼에서 peak와 파장 간격 감소를 다시 확인한다.
- 원시 후보 bundle 2개와 멤버 **24개**의 archive/member SHA256 및 길이를 검증했다. 원 데이터 숫자를 반올림하지 않았다.
- baseline은 b9105dd의 V1–V9 record canonical SHA256로 고정한다. 과거 PASS의 형상·모드 범위를 유지한다. hash는 입력/파일 동일성이지 외부 진실성·제품 인증의 서명이 아니다.

## 테스트·CI

전체 Python 회귀 **163개 PASS**. 이 중 이번 작업의 신규 물리-evidence 로직은 **23개**이며, V2/V3 완전 합성 package 평가, 결측·부호·정규화 불일치·hash 변조·잘못된 장 대응·독립 스윕의 결합 입력 변경·V6 오차 조작·V7 branch 단절 등의 회귀를 포함한다. 테스트 fixture PASS는 실제 COMSOL 외부 비교의 PASS가 아니다. 새 모듈 compileall 및 workflow YAML 구문도 확인했다.

새 `COMSOL physical evidence gate` workflow는 `evidence-integrity`와 `physical-acceptance`를 분리한다. 후자는 `--require-all`을 실행하므로 현재 자료에서는 실패해야 올바르다. 실패해도 gate 결과를 Actions artifact에 보존한다. 기존 software workflow의 성공을 물리 완료로 승격시키지 않는다. 이 문서는 로컬 실행 결과이며 원격 CI 결과를 선제적으로 성공했다고 표시하지 않는다.

- [최신 판정 JSON](PHYSICS_GATE_20261121.json)
- [전체 Python 로그](physics_20261121/python_tests.txt)
- [strict gate 로그](physics_20261121/strict_gate.txt)
- [실행 환경·현재 소스 hash](physics_20261121/environment.json)
- [결측 입력 manifest](physics_20261121/external_manifest.json)
- [수용 정책·자료 계약·실행 명령](PHYSICS_ACCEPTANCE_POLICY.md)

## 남은 필수 작업

1. COMSOL Application Library 원본 microstructured_optical_fiber_parameters.txt와 동일 버전 모델/장 확보.
2. V3 full-precision bent complex export 및 실제 rAverage/r0, 입력·PML·정식화 대응 확보.
3. V6 높은 cutoff/메시 및 독립 PML/window/ring 수렴, HOM 장·cutoff bracket, primary curve/불확도 검증.
4. V7 공진 주변 탐색/혼성 모드 식별 개선과 같은 branch의 peak·독립 수렴 검증.

이번 네 단계는 근거 저장과 자동 평가 경로를 갖춘 결과다. 물리 정확도·ESM·HCF 극소손실·실제 제품 인증을 완료한 결과로 해석하지 않는다.
