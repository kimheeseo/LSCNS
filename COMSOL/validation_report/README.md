# COMSOL 검증 보고서 모음

## YOFC G.654.E — Step-index 근사 모델 검증 (2026-11-03)

![계산 결과 대시보드](./YOFC_G.654.E/assets/dashboard.svg)
- **[상세 검증 보고서 (Markdown)](./YOFC_G.654.E/VALIDATION_REPORT.md)** — 그림별 해석과 YOFC 공개 사양 비교.

**해석 범위:** 파장 1550 nm, Step-index 코어 반경 5.4 µm, Δn=0.004, 계산 영역 반경 40 µm, radial FEM 750 elements. MFD·Aeff는 공개된 대표값에 맞추어 입력을 조정한 것이므로 **실제 YOFC 제품의 독립 검증은 아닙니다.**

**핵심 결과:** n_eff 1.446099217, MFD_PII 11.8340 µm, Aeff 110.3678 µm², D1550 20.1031 ps/(nm·km), D slope 0.059088 ps/(nm²·km).
