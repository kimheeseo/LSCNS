# COMSOL Step-index 직접 웹 FEM 재현 (2026-11-04)

- 공식 COMSOL 자료: https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html
- 코어 반경 8 µm, 클래딩 반경 40 µm, λ1550 nm, n_core=1.4457, n_clad=1.4378
- 공개 COMSOL 기본 모드 neff **1.4444** (반올림 4자리)
- 실제 LSCNS Scalar FEM (750 요소) neff **1.444401819379**
- 상대오차 약 **0.000125961%** → 요청 5% 미만 PASS
- 독립 LP01 Bessel 해석해 **1.444401844409**와 비교 및 200~1800 요소 수렴 검증
- **테스트용 치환 코드가 아니라**, 현재 정식 브라우저 솔버의 opt-in 고정 굴절률 설계 예제로 계산. 공개 neff를 솔버 계산에 넣거나 피팅하지 않음.
- 실행: node COMSOL/validation_report/comsol_step_index_benchmark/run.mjs
- [웹 검증 방법](WEB_PRESET_VALIDATION.md)

COMSOL 원본은 벡터 모드이며 이 구현은 Scalar LP01. 공식 neff가 반올림 4자리로 제공되어 실제 COMSOL 고정밀 오차를 주장할 수 없고, 일반 Step-index 예제이지 제조사 G.654.E 실측 RIP 검증은 아닙니다. 고정 굴절률로 계산한 D는 실제 재료 분산이 아닙니다.
