# M4 — 3D Wave Optics (현재 저장소 복구 상태)

**중요:** 이전 대화에서 3D Nédélec FEM / Cartesian PML / mode ports / S-matrix 구현과 104개 테스트 통과가 보고되었으나, COMSOL GitHub 폴더에는 해당 완전한 M4 소스 및 실행 증빙이 없었습니다. 본 폴더는 그 코드를 복구했다고 주장하지 않습니다.

대신 `src/m4_cavity.py`에 독립적으로 재구현한 **3D 사면체 Nédélec PEC cavity eigenmode** 솔버를 제공합니다. GitHub Actions 수치 검증 후 `M4_VALIDATION_REPORT.md`, `screenshots/`에 계산 결과를 생성합니다. 이는 전체 M4 PML/포트/S 검증을 대체하지 않습니다.

실행: `python src/m4_cavity.py --out screenshots` (NumPy, SciPy, Matplotlib). 수치 결과는 교육 및 수치해석 검증용이며 COMSOL과 독립 직접 비교는 수행하지 않았습니다.
