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
| 실제 웹 | STEP6 탭/예시/버전 입력/무효화/모바일/HCF·vector navigation | GitHub Actions 실행 후 별도 CI 결과 기록 |

실제 수치 실행시간 210.254 s (현 환경). [원 결과](step6_v088/validation.json) · [Python 전체 로그](step6_v088/python_tests.txt) · [실행/공식/출처](../universal_silica_v01/fiber_research_v1/STEP6_V088.md) · [한계/다음 제안](STEP6_V088_LIMITATIONS_20261120.md).

소스 확인: COMSOL 6.4 bend/stress/directional coupler/fiber coupling 공식 문서, Ansys mode overlap 공식 해설, Matsushima & Shimobaba angular spectrum 저자 원문 메타데이터, ITU G.650.2 정의, Folkenberg PCF primary paper. 개별 URL과 적용 범위는 실행 문서의 근거 표를 따른다. 미확인 재료 상수는 TODO다.
