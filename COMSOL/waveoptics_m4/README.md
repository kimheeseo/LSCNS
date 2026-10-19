# M4 — 3D PEC 공동 FEM

`src/m4_cavity.py`는 사면체 1차 Nédélec PEC 공동 고유모드 솔버입니다. 기본 검증은 (4,3,6), (8,6,12), (12,9,18)의 세 메시에서 실행하며, 최세밀 TE101 주파수 오차 0.1% 미만을 요구합니다.

별도의 전체 3D 주파수 영역 PML·직사각형 포트·S 행렬 소스는 [보관 자료](../reports/M4_3D_2026-10-11/)의 ZIP에 들어 있습니다. 이 공동 솔버와 혼동하지 마십시오.

저장소 루트에서 `python COMSOL/run_m4_m6_validation.py`로 공동 검증을 재현합니다. 전체 3D 핵심·추가 17개 시험은 `python COMSOL/research/run_regressions.py --only M4_full`로 실행합니다.

[연구용 검증 안내](../research/README.md) · [공동 보고서](M4_VALIDATION_REPORT.md)
