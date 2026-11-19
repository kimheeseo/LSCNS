# STEP6 v0.8.8 검증표 — 2026-11-20

코드·웹·예시를 기존 자산 위에 추가했다. 아래 수치는 합성 수치해석 파이프라인 시험이며 COMSOL 동등성 또는 상용 광섬유 성능 인증이 아니다.

| 단계 | 실행 및 결과 | 판정 / 한계 |
|---|---|---|
| 회귀 | Python 124 tests / Node 8 tests PASS | 기존 동작과 새 독립 공식 시험 포함 |
| 6a 굽힘 | R=100,20 mm × 방향 0,90°: 4/4 공통 probe 부분공간 추적, reference rank=2 | PIPELINE PASS; 작은 창의 높은 기저 누설, 정량 bend loss 미인증 |
| 6b PM | 합성 재료 cooldown, 공간 overlap=0.997200584, B=2.67249987e-05, L_B=0.057998132 m | SYNTHETIC ONLY; PANDA 상수·실측/COMSOL reference 미확인 |
| 6b 무응력 | ΔT=0, max stress=0 Pa | 영 냉각 항등식 PASS |
| 6c MCF | 2코어/3코어 각각 실제 FEM 6모드 반환. unique even/odd Jones-axis overlap=0.987350637; Reκ=204.893063 m⁻¹, Lc=0.00766642024 m | CANDIDATE; 동일 원점 대칭 2코어 한정. 완전 guided spectrum 및 편광 퇴화 수렴은 미확인 |
| 6d 독립 해 | Gaussian size mismatch/offset/angle/gap, orthogonal polarization 및 normal Fresnel 시험 PASS | 요청한 Et overlap/ASM 근사 검증; full E/H interface 아님 |
| 6d SMF↔HCF | 실제 FE grid → Et overlap 계산·저장. HCF edge peak fraction=0.367393 | PIPELINE ONLY; HCF 공간 모드/창 검사가 미흡해 물리 접속손실 검증 FAIL. 수치 η/손실은 후보 raw 값이며 제품 예측값 아님 |
| 6e 다차원 | 3코어 trench pitch=6,7 µm × width=.3,.5 µm, 4/4 실제 FEM. Pareto case=[0] 및 목적별 최적 입력 저장 | FINITE GRID PASS; 후보 모드 집합 목적이며 특정 FM branch 최적/전역 최적 아님 |
| 6e 예외 | failure/nonfinite 제외, feasibility, checkpoint resume/hash mismatch, Pareto 및 best input 시험 | PASS |
| 6f | signed Δn=a+bλ → Δng=a 독립 항등식 PASS. 기존 PMD 시연용 유지 | Uniform uncoupled DGD 경로만 추가, random PMD 미구현 |
| PCF V2 | 공식 neff 표 확인, 정확한 외부 parameter txt 미확인 | BLOCKED, 통과 아님 |
| PCF V6 | 문헌 7개 geometry/Table I와 Eq1 확인 | FEM cutoff NOT VALIDATED, fit은 참고 데이터 |
| 실제 웹 | STEP6 탭/예시/버전 입력/무효화/모바일/HCF·vector navigation 및 전역 DOM ID 중복 검사 | PASS. browser page_errors=[], 모바일 overflow=false. 기존 굽힘 패널 ID 충돌 수정 후 전체 워크벤치 회귀 PASS |

실제 수치 실행시간 210.254 s (현 환경). [원 결과](step6_v088/validation.json) · [Python 전체 로그](step6_v088/python_tests.txt) · [실행/공식/출처](../universal_silica_v01/fiber_research_v1/STEP6_V088.md) · [한계/다음 제안](STEP6_V088_LIMITATIONS_20261120.md).

소스 확인: COMSOL 6.4 bend/stress/directional coupler/fiber coupling 공식 문서, Ansys mode overlap 공식 해설, Matsushima & Shimobaba angular spectrum 저자 원문 메타데이터, ITU G.650.2 정의, Folkenberg PCF primary paper. 개별 URL과 적용 범위는 실행 문서의 근거 표를 따른다. 미확인 재료 상수는 TODO다.

수정 이력: 신규 STEP6 DOM ID가 legacy STEP6a의 `r6-open`, `r6-status`와 겹친 회귀를 확인해 신규 패널을 `r688-*` 이름으로 분리했다. 전역 중복 ID 검사와 기존 full-workbench GUI 시험을 모두 다시 통과했다. Python 수치 코드는 이 수정에서 변경하지 않았다.

## GitHub 최종 확인

수치 코드 및 UI 검증 커밋: `a3e08e9b7dbb98f6d64e97ab95c4d764daf872b1`. GitHub Actions **8/8 성공**: Python 124 tests, Node 8 tests, 독립 STEP3 V1/V4/V5, STEP4·5 및 STEP6 실제 계산, 기존 25개 프리셋 Gmsh 메시, 기존 full-workbench/HCF/vector UI와 STEP3/4/5/6 browser, GitHub Pages 배포. 신규 STEP6는 page_errors=[], mobile_overflow=false.

[CI 요약·실행 링크](step6_v088/ci_summary.json) · [CI STEP6 실제 수치](step6_v088/ci_numerical_validation.json) · [신규 브라우저 결과](step6_v088/browser_check.json). 각 검증은 위 표의 물리적 적용 범위를 확장하지 않는다. 후속 문서 커밋은 계산·웹 코드 변경 없이 이 증거만 추가한다.
