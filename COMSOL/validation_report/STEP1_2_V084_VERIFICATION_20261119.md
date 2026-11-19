# STEP 1·2 v0.8.4 — 실행 검증표 (2026-11-19 KST)

사용자 승인: STEP 0 보고서 GitHub 업로드 및 STEP 1·2 진행 요청. STEP 3 정식화/고유값 코드는 변경하지 않았다.

## 재현 환경

Python 3.12.14, numpy 2.3.5, scipy 1.17.0, gmsh 4.15.2, scikit-fem 12.0.2, jsonschema 4.26.0.
로컬 실제 실행 결과이며 COMSOL .mph 실행 또는 제품 인증 결과가 아니다. GitHub Actions의 Python 3.11 재실행은 별도 환경이다.

| 검사 | 결과 | 근거·범위 |
| --- | --- | --- |
| 기존 테스트 baseline | 67/67 PASS | 수정 전 실행 |
| 전체 Python 테스트 | 83/83 PASS | [실행 로그](step12_v084/python_tests.txt), 기존 벡터/굽힘/PM 경로 포함. 새 정확도 인증 아님 |
| 25개 JSON 계약·왕복 | PASS | JSON Schema + Python + 저장/불러오기 |
| 잘못된 입력·간섭 | PASS | 타입/NaN/Boolean, domain clipping, ellipse overlap, ring MCF chord, trench·nested·PBGF mismatch 등 |
| y-major 타원 | PASS | OCC 생성 성공, 독립 πab CAD 면적 및 x/y 반축 검사 |
| 기본 프리셋 실제 기하 생성 | 25/25 PASS | [전체 결과](step12_v084/validation_summary.json) |
| CAD 대비 선형 삼각형 재료 면적 | 25/25 PASS | 최악 상대 차이 0.303366%, 명시한 검사 정책 2% 미만 |
| 원형 annulus 벽 법선 교차 | 75개 벽 PASS_SAMPLED | 각 기본 벽에서 24개 각도, 최소 횡단 요소수와 coverage≥0.99. 전 경계 보증 아님 |
| 기본 벌크 파장 기준 | **25/25 FAIL_POLICY** | 기존 boundary_refined 호환 정책의 제약을 숨기지 않음. 새 wavelength 옵션 사용 필요 |
| 엄격 파장 모드·3개 실제 메시 | PASS_POLICY | 별도 소형 Bragg fixture, 아래 표 |
| 전기/광학 물성 추정 금지 | 유지 | Ge/F 조성 계수 TODO, 기존 상수·예시값에 출처/한계 표시 |
| 새 웹 JSON import 검사 | Node 2/2 PASS, syntax PASS | 비유한 좌표/경계 누락·입력 비교 |
| 로컬 브라우저 화면 통합 | 미완료 | Chromium 실행이 socket 권한 제한으로 중단됨. 실행 성공으로 표기하지 않음 |
| GitHub 브라우저 통합 | **PASS (CI 실제 실행)** | 실제 Python mesh_preview import·입력 변경 무효화·390px 패널 검사 스크립트 포함 |
| 광학 모드/손실 수렴 | **미검증** | 본 단계는 기하/메시. neff/Im(neff) 인증 안 함 |

## 실제 3단계 기하 refinement

동일 소형 Bragg fixture의 벌크·벽·곡률·파장 해상도를 모두 변화시켰다. 기하 면적 오차가 모든 재료에 대해 단조 감소한다고 가정하지 않는다. 광학 모드 수렴 결과로 해석하면 안 된다.

| refinement factor | triangles | high-layer 면적 상대차이 | 파장 정책 | 벽 표본 |
| --- | --- | --- | --- | --- |
| 1.0 | 15,047 | 8.2747681e-06 | PASS_POLICY | PASS_SAMPLED |
| 1.3 | 24,794 | 1.2179273e-05 | PASS_POLICY | PASS_SAMPLED |
| 1.6 | 35,709 | 1.1628907e-05 | PASS_POLICY | PASS_SAMPLED |

## 재실행

    cd COMSOL/universal_silica_v01/fiber_research_v1
    python -m pip install -r requirements.txt
    python -m unittest discover -s tests -v
    python validate_step12.py --out validation_step12.json
    python cli.py mesh-study benchmarks/step12_bragg_mesh_input.json --out out_mesh_study

웹 패널 검사 명령과 fixture 요구사항은 web/tests/browser_research_design_mesh_v084.cjs를 참고한다. 이번 보고서에 없는 COMSOL 오차·HCF dB/km·Panda modal B의 PASS는 주장하지 않는다.

## GitHub CI 최종 확인

[검증 실행 #57 이후 수정본](https://github.com/kimheeseo/LSCNS/actions/runs/44448892983): **전체 성공**. Python 3.11 환경의 83개 테스트·25개 JSON·25개 실제 메시 생성·실제 Python 메시 가져오기·입력 변경 무효화·390px 신규 패널 검사를 통과했다. 브라우저 페이지 오류 0개, 신규 패널 가로 넘침 없음. [기존 COMSOL형 GUI 회귀](https://github.com/kimheeseo/LSCNS/actions/runs/44448484518)와 [HCF/Full-vector 프리셋 UI](https://github.com/kimheeseo/LSCNS/actions/runs/44448484505)도 통과했다. 범용 Web Smoke의 별도 실행은 보고서 갱신 시점에 진행 중이었다.

로컬 Chromium의 socket 제한은 GitHub 브라우저 실행으로 보완했다. 이 성공은 광학 손실·COMSOL 수치 정확도 인증과는 별개다. [기계 판독 요약](step12_v084/ci_summary.json).
