# COMSOL 공개 Step-index neff 웹 직접 검증 (2026-11-04)

**목표:** 5% 미만 neff 상대오차. **결과:** 약 **0.000126%**. 공식 COMSOL 결과는 4자리 반올림값으로 공개되므로 실제 COMSOL 고정밀 수치와의 정확한 오차를 의미하지 않음.

## 사용법

1. [기존 Universal Silica Fiber Lab](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/) 왼쪽 **설계 예제**에서 **COMSOL 공개 Step-index · neff 검증** 선택.
2. 자동으로 **1550 nm / 코어 반경 8 µm / 클래딩 경계 40 µm / n_clad 1.4378 / Δn 0.0079 (n_core 1.4457)** 설정.
3. 브라우저가 실제 **Scalar Radial P1 FEM**을 실행해 neff 및 모드장을 구함. COMSOL 공개값 1.4444는 **계산 종료 후 비교에만** 이용하며 고유값 솔버에는 넣지 않음.
4. 결과 카드에 기준값, 현재 neff, 절대오차, 상대오차, **5% PASS/FAIL** 및 공개 반올림 정밀도 비교 표시.
5. 입력을 바꾸면 비교가 무효화됨. 고정 굴절률은 해당 파장에서만 알려진 값이므로 소재 분산·기울기 검증은 하지 않음.
6. G.654.E 등 기존 실리카 Sellmeier 프리셋은 원래 물리 모델 그대로 유지.

## 공식 입력값 및 결과

[COMSOL Step-Index Fiber 공식 문서](https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html): λ 1.55 µm, n_core 1.4457, n_clad 1.4378, a_core 8 µm, a_clad 40 µm, neff **1.4444** (4자리 공개).

| 검증항목 | 값 |
|---|---:|
| COMSOL 공개 neff (4자리) | 1.4444 |
| LSCNS 현재 Scalar FEM (750 요소) | 1.444401819379 |
| 절대차 | 1.819379×10⁻⁶ |
| 상대오차 | 0.000125961% |
| 독립 LP01 Bessel 값 | 1.444401844409 |
| COMSOL 반올림(±0.00005) 일치 | 예 |
| 상대오차 <5% | PASS |

**검증 제한:** COMSOL은 벡터 Maxwell 해석이고 본 브라우저 솔버는 약유도 Scalar LP01. COMSOL 공개 neff는 4자리이므로 1e-6 수준의 정량 COMSOL 오차를 선언하지 말 것. 이 COMSOL 공개 모델은 **G.654.E 제품별 RIP 검증과 다른 작업**.

### 자동 테스트
- [직접 원본 솔버 테스트](run.mjs) 및 [GitHub Actions](https://github.com/kimheeseo/LSCNS/actions/workflows/comsol-step-index-benchmark.yml).
- 기존 G.654.E Sellmeier 프리셋 회귀 테스트.
- HCF·Full-vector 메뉴를 포함한 실시간 Chromium UI 테스트에서 COMSOL 프리셋 및 수정 후 비교 무효화 확인.
