# STEP 6a — 굽힘 해석 설계·구현·검증표 (v0.8.1)

일자: 2026-11-18 (KST).
범위: STEP 6a ONLY. STEP 6b~6f, STEP 7 V1~V9 전체 UI, STEP 8 전체 리팩토링은 미착수.

## 1. 물리식 및 설계
- 좌표 단위 x,y [µm], 굽힘 반경 R [mm], 방향 θ [deg CCW, +x가 굽힘 바깥쪽].
- x_out = x cosθ + y sinθ [µm].
- n_eq(x,y)=n(x,y)[1+x_out/R_eff], R_eff[µm]=1000 R[mm] f_literature.
- FEM 유전율 텐서 처리: ε_eq(x,y)=ε_material(x,y)[1+x_out/R_eff]²를 복소 PML 텐서 적용 **전에** 구성.
- f_literature 기본값 1 (응력 보정 없는 기하학 모델). 별도 보정은 인용 가능한 출처 URL, 유효계수와 적용 조건을 사용자가 모두 입력해야 함. 모든 섬유에 적용할 보편 상수는 존재하지 않는다고 가정.
- R_eff가 외곽 해석 반경 이하인 입력을 거부하고 R_eff/외곽 반경 비를 진단값으로 보존.
- 유출 복소 neff에서 α_total[dB/m]=(20/ln10)(2π/λ[m])Im(neff). bend radiation-only로 명명하지 않음. Δα_total=α_bent−α_straight. 1회 굽힘 총손실=α_total×2πR[mm]×10⁻³ [dB].
- 직선과 모든 굽힘 방향/반경에서 동일한 물리 메시를 재사용. 낮은 곡률부터(큰 R) 계산하고 복소 Et overlap으로 모드가 바뀌지 않게 일대일 배정. mismatch/불안정 시 결과 null.
- 설계 스키마 v1.0.0 호환 유지, 엔진 v0.8.1. 입력 JSON, sweep 설정, 근거 URL과 그에 대응하는 SHA-256, UTC 시각을 결과 JSON에 기록. 원본 설계 수정 시 웹 곡선 무효화.

## 2. 구현 경로
- Python bending.py: conformal factor, 단위 변환, 사용자 문헌 R_eff, 반경/방향 sweep, mode overlap, 오류 차단.
- Python vector_fem.py: 기존 FEM/PML 유전체 조립 전에 공간적 굽힘 보정; bent case engine_version 0.8.1.
- Python cli.py: bend 명령 — radii-mm, directions-deg, reff-multiplier, reff-source-url, reff-source-note, modes, min-overlap.
- Web research_step6a_v081.mjs: 기존 설계기 내부 STEP6a 결과 뷰어, 방향별 반경-총손실/변화/회당 dB 그래프, 계산 근거, 문헌 링크, 설계 일치/불일치 표시.
- Test test_step6a_bending.py: 수식/단위/방향/재료 epsilon²/옵션 근거/해시/모드 순서 변화/유효경계 검사.
- Test test_vector_smoke.py: 실제 소형 Gmsh + scikit-fem mixed vector bend case 1개 스모크.
- Smoke JSON benchmarks/step6a_step_index_smoke_input.json: TEST ONLY (메시/PML 수렴 및 COMSOL V3 비교 증거 아님).

## 3. 검증표

| 검증 | 기준 | 상태 |
| --- | --- | --- |
| S6a-01 등각 좌표 | 100µm/10mm 방향 0° ⇒ 1.01, 90° ⇒ x효과 0 | PASS |
| S6a-02 FEM epsilon | eps_eq=eps*(1+x_out/R_eff)², PML 전 적용 | PASS (단위) |
| S6a-03 R_eff | 인용계수·source URL·조건 누락 시 거부 | PASS |
| S6a-04 단위/범위 | mm↔µm, NaN/0/음수, R_eff <= outer 반경 거부 | PASS |
| S6a-05 손실 계산 | α[dB/m], Δα, signed dB/turn | PASS |
| S6a-06 모드 추적 | x/y 모드 순서 뒤집기 대응, 불일치 시 손실 값 null | PASS |
| S6a-07 방향/반경 | 2방향 × 2반경, 고정 메시, 큰 반경 우선 | PASS (mock) |
| S6a-08 입력 해시 | 방향, 반경, 출처 보정 변경 시 SHA-256 변경 | PASS |
| S6a-09 진짜 FEM 스모크 | Gmsh mesh + 벤드 QEP candidate 산출 | PASS (smoke only) |
| V3 COMSOL bend | 동일 모델/모드/COMSOL CSV에 근거한 Re/Im(neff) 오차 | **EVIDENCE_REQUIRED / 미통과** |
| V4 PML 독립성 | PML 두께/강도/도메인 변화 시 Im(neff) 수렴 | **NOT PASS** |
| V5 메시 수렴 | h vs Re/Im(neff), loss stability | **NOT PASS** |

검증 CI 실행: https://github.com/kimheeseo/LSCNS/actions/runs/44181794294 .
기존 STEP 4 공식/테스트 및 브라우저 JS node --check도 CI에 포함. 수식 PASS는 실제 굽힘 손실 정확도 PASS와 다름.

## 4. 사용 및 데이터

    cd COMSOL/universal_silica_v01/fiber_research_v1
    python cli.py validate benchmarks/step6a_step_index_smoke_input.json
    python cli.py bend benchmarks/step6a_step_index_smoke_input.json --radii-mm 50 20 10 --directions-deg 0 90 --modes 2 --out out_bend

결과: out_bend/bend_sweep_result.json 을 기존 웹 STEP 6a 패널에서 업로드.
예시 반경 50/20/10 mm는 명령어 예제일 뿐 적정한 케이블 굽힘반경이나 실측 결과가 아님.

R_eff를 문헌에서 확인한 경우에만 다음 추가 옵션 지정:
- --reff-multiplier [source-dependent physical coefficient]
- --reff-source-url [actual literature URL]
- --reff-source-note [applicability: silica type, lambda, geometry, temperature]

## 5. 단계별 승인 게이트
V3/V4/V5는 신뢰 가능한 수치로 통과하지 않았으므로 현재 굽힘 손실을 검증된 결과라고 표시하면 안 된다.
다음 STEP 6b (열응력·광탄성) 설계/구현, STEP7 검증 탭, STEP8 구조 개편은 사용자 추가 승인 전 착수하지 않는다.

## 출처
- COMSOL Step-Index Fiber Bend v6.4 (2D axisymmetric model, PML): https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.step_index_fiber_bend/step_index_fiber_bend.html
- COMSOL conference bend conformal and R_eff: https://www.comsol.com/paper/study-of-bending-losses-in-optical-fibers-using-comsol-multiphysics-61111
- Optiwave stress-optic R_eff (silica example; no universal constant): https://optiwave.wiki/kb/optical_fiber_bending_loss_with_stress_effect/
