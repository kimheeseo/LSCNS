# COMSOL validation report index

## YOFC G.654.E — re-executed with current Web FEM v0.2.1 (2026-11-03)

[![YOFC zero dispersion / slope diagnostics](./YOFC_G.654.E/assets/dispersion_slopes.svg)](./YOFC_G.654.E/VISUAL_REPORT.html)

- **[시각화 HTML 보고서](./YOFC_G.654.E/VISUAL_REPORT.html)** — 1550 nm 분석, 광섬유 단면, 굴절률 프로파일, λ₀·S₀·S₁₅₅₀ 및 메시 수렴.
- **[상세 검증 보고서](./YOFC_G.654.E/VALIDATION_REPORT.md)** — 제조사 기준값과 차이, S₀/S₁₅₅₀ 정의 및 해석 한계.
- **[수정된 Jupyter Notebook](./YOFC_G.654.E/YOFC_G.654.E_validation.ipynb)** — 기존 SciPy FEM과 현 웹 `solver.mjs` 소스를 함께 포함한 재현 노트북.
- **[실제 웹 솔버 실행값(JSON)](./YOFC_G.654.E/web_solver_rerun.json)** — current engine 0.2.1 결과, 입력조건, 기울기, 메시 수렴, 파장 스윕.
- [CSV 및 기타 산출물](./YOFC_G.654.E/)

**현재 YOFC 가상 Step-index 모델 결과:** MFD 11.8340 µm, Aeff 110.3678 µm², D@1550 20.1031 ps/(nm·km), S₁₅₅₀ 0.059089 ps/(nm²·km), λ₀ 1275.457 nm, S₀ 0.093979 ps/(nm²·km). **S₀에는 YOFC 1550 nm 상한 0.070을 적용하지 않습니다.**

**검증 범위:** 굴절률의 실제 측정 프로파일이 아닌 대표 MFD/Aeff를 맞춘 근사 Step-index 모델입니다. **실제 YOFC 제품 인증이나 COMSOL Wave Optics 실행 결과가 아닙니다.**
