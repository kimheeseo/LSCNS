# Python Wave Optics — M1–M6

- [M1 slab scalar modes](waveoptics_m1/)
- [M2 vector modes](waveoptics_m2/)
- [M3 TE/TM scattering + PML + S matrix](waveoptics_m3/) — [검증 보고서](waveoptics_m3/M3_VALIDATION_REPORT.md), [스크린샷](waveoptics_m3/screenshots/)
- [M4 3D Nédélec PEC cavity (verified recovery subset)](waveoptics_m4/) — **과거 M4 PML/port/S 전체 원본 GitHub 미보존**, 제한적 3D FEM 검증
- [M5 material dispersion + spectral S](waveoptics_m5/)
- [M6 constrained broadband inverse design](waveoptics_m6/)

M4/M5/M6는 `python COMSOL/run_m4_m6_validation.py`로 독립 수치 검증 및 PNG/CSV/JSON/MD 보고서 생성. GitHub Actions 워크플로: `.github/workflows/waveoptics-m4-m6-validation.yml`.

검증 수준을 구별하십시오. M1–M3의 원래 실행 증거는 각 폴더에 있습니다. M4의 *이전 대화상 작업*과 *본 저장소에 다시 구현된 제한적 솔버*는 동일한 범위로 취급하지 않습니다. M5/M6는 1D 독립 오라클이며 상용 COMSOL 검증 또는 실제 상품 스펙 비교가 아닙니다.
