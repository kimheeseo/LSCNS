## STEP 5 v0.8.3 — Photonic Crystal Bloch-Floquet

- [STEP5 설계·검증표](STEP5_V083_DESIGN_VERIFICATION_20261118.md)
- [STEP5 별도 한계 보고서](STEP5_V083_LIMITATIONS_20261118.md)
- [STEP5 동일 웹 설계 UI](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#research-step5-v083)
- Python Floquet 스칼라 예제의 표본 gap 후보 0개. V2/V4/V5/V6 정량 검증은 아직 미통과.


# COMSOL validation report index

## 2026-11-18 — STEP 6b PM 열응력·광탄성 v0.8.2 (단일 단계)

- [설계·구현 및 58개 스모크/단위 테스트 검증표](./STEP6B_V082_DESIGN_VERIFICATION_20261118.md) — generalized plane strain, user-sourced E/ν/α/p11/p12, photoelastic Q tensor, Panda 별도 stress_rod, 임시 벡터 모드 B.
- [별도 STEP 6b 한계 보고](./STEP6B_V082_LIMITATIONS_20261118.md) — 실제 도핑/냉각/광탄성 자료 TODO, COMSOL/neff PML 정밀도 미확인.
- [Python stress_optic.py/pm_cli.py](../universal_silica_v01/fiber_research_v1/) · [웹 STEP6b 결과 뷰어](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#research-step6b-v082).
- [58개 Python 테스트/새 웹 모듈 syntax 통과 CI](https://github.com/kimheeseo/LSCNS/actions/runs/44185058249).
- **중요:** 재료 평균 B는 optical modal B가 아님. COMSOL 비교/V4·V5 정밀도 검증 미통과. STEP6c~6f 및 STEP7/8 전체는 승인 전 착수 금지.

## 2026-11-18 — STEP 6a 굽힘 해석 v0.8.1 (단일 단계)

- **[STEP6a 설계·구현·42개 CI 테스트 검증표](./STEP6A_V081_DESIGN_VERIFICATION_20261118.md)** — 2D 등가 굴절률, 반경/방향, optional sourced R_eff, 동일 메시 overlap 추적.
- **[STEP6a 한계 보고서](./STEP6A_V081_LIMITATIONS_20261118.md)** — V3 COMSOL axisymmetric와 차이, PML/mesh 수렴 미완, 응력광학 계수 근거 TODO.
- [Python 굽힘 계산/CLI](../universal_silica_v01/fiber_research_v1/) · [웹 결과 뷰어](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#research-step6a-v081).
- [CI 42개 테스트 및 JavaScript 점검 성공](https://github.com/kimheeseo/LSCNS/actions/runs/44181794294).
- **V3 = EVIDENCE_REQUIRED, V4/V5 = NOT PASS.** STEP6b~6f 및 STEP7/8 전체 구현은 승인 전 보류.

## 2026-11-18 — STEP 4 계산 물리량 v0.8.0 (한 단계만)

- **[설계 / 구현 / 29개 테스트 검증표](./STEP4_V080_DESIGN_VERIFICATION_20261118.md)** — signed α[dB/m], Et-only Aeff/MFD, overlap-gated D/S/λ0, γ, HCF HOMER·ARROW 등.
- **[별도 한계 보고서](./STEP4_V080_LIMITATIONS_20261118.md)** — 물성 TODO, FEM/PML·surface-mode·stress 등 미검증 상세.
- [Python STEP4 수식·CLI](../universal_silica_v01/fiber_research_v1/) · [기존 웹 STEP4 뷰어](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#research-step4-v080).
- [통과한 CI 실행 (29 테스트 + 정적 JS 검사)](https://github.com/kimheeseo/LSCNS/actions/runs/44138275833).
- **수용 게이트:** STEP3 V1/V4/V5 물리 수렴 증거 미제출. **STEP5 미착수·승인 전 진행 금지.**

## 2026-11-18 — STEP 0 감사 승인 및 STEP 1–3 연구용 광섬유 설계/솔버 확장

- **[STEP 0 승인·STEP 1–3 검증 수용 기준](./STEP0_APPROVAL_STEP1_3_ACCEPTANCE_20261118.md)** — 기존 감사 결과 보존과 향후 정확도 기준.
- **[STEP 1–3 실제 구현 상태·CI 검증 기록](./STEP1_3_IMPLEMENTATION_STATUS_20261118.md)** — 25개 JSON 프리셋, Gmsh 메시, mixed HCurl/H1 풀벡터 Python 스모크 테스트 9/9 성공, 수치·물리적 미검증 항목 명시.
- [Python 연구용 독립 엔진](../universal_silica_v01/fiber_research_v1/) / [기존 웹의 독립 결과 뷰어](../universal_silica_v01/web/) — 기존 M1–M6 스칼라/Yee 기능 및 프로젝트 파일 경로 유지.
- [GitHub Actions 실행](https://github.com/kimheeseo/LSCNS/actions/runs/44134039197) — gmsh/scikit-fem/CLI 스모크 테스트. **광학적 정밀도나 PML 누설손실 검증 합격을 의미하지 않음.**

**범위:** 미구현 PBGF Bloch bandgap, 정확한 Kagome/conjoined 형상, PM 응력광학, C6v 벡터 대칭 조건 및 Ge/F 도핑 Sellmeier 계수는 검증 전까지 TODO/근사로 분리합니다. PCF ESM 지도는 참고 문헌의 d/Λ 영역만 보여주는 정성적 도식입니다.


## YOFC G.654.E — re-executed with current Web FEM v0.2.1 (2026-11-03)

[![YOFC zero dispersion / slope diagnostics](./YOFC_G.654.E/assets/dispersion_slopes.svg)](./YOFC_G.654.E/VISUAL_REPORT.html)

- **[시각화 HTML 보고서](./YOFC_G.654.E/VISUAL_REPORT.html)** — 1550 nm 분석, 광섬유 단면, 굴절률 프로파일, λ₀·S₀·S₁₅₅₀ 및 메시 수렴.
- **[상세 검증 보고서](./YOFC_G.654.E/VALIDATION_REPORT.md)** — 제조사 기준값과 차이, S₀/S₁₅₅₀ 정의 및 해석 한계.
- **[수정된 Jupyter Notebook](./YOFC_G.654.E/YOFC_G.654.E_validation.ipynb)** — 기존 SciPy FEM과 현 웹 `solver.mjs` 소스를 함께 포함한 재현 노트북.
- **[실제 웹 솔버 실행값(JSON)](./YOFC_G.654.E/web_solver_rerun.json)** — current engine 0.2.1 결과, 입력조건, 기울기, 메시 수렴, 파장 스윕.
- [CSV 및 기타 산출물](./YOFC_G.654.E/)

**현재 YOFC 가상 Step-index 모델 결과:** MFD 11.8340 µm, Aeff 110.3678 µm², D@1550 20.1031 ps/(nm·km), S₁₅₅₀ 0.059089 ps/(nm²·km), λ₀ 1275.457 nm, S₀ 0.093979 ps/(nm²·km). **S₀에는 YOFC 1550 nm 상한 0.070을 적용하지 않습니다.**

**검증 범위:** 굴절률의 실제 측정 프로파일이 아닌 대표 MFD/Aeff를 맞춘 근사 Step-index 모델입니다. **실제 YOFC 제품 인증이나 COMSOL Wave Optics 실행 결과가 아닙니다.**
