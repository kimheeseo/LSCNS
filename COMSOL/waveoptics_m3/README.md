# M3 — 2D Wave Optics TE/TM / PML / S-Parameters

기존 검증 코드와 계산 자료를 보존한 M3 아카이브.

**검증:** 87개 pytest 통과(기존 실행). P1 삼각형 FEM, 2D TE(Ez)/TM(Hz), Cartesian PML, modal ports, 복소 S11/S21 및 원통 산란 해석.

- [M3 상세 검증 보고서](M3_VALIDATION_REPORT.md)
- [정식화](M3_FORMULATION.md)
- [스크린샷 폴더](screenshots/)
- [원시 CSV/NPZ/로그](results/)
- [수치 검증 코드](validate_m3.py)

M3는 2D 산란 모델이며 3D full-vector PML 또는 광섬유 상용 제품의 검증은 포함하지 않습니다.
