# STEP 6b — PM 섬유 열응력·광탄성 (v0.8.2) 설계·구현·검증표

일자: 2026-11-18 (KST). **이번 승인 범위는 STEP 6b ONLY**. STEP 6c–6f와 STEP7·8 전체 기능은 미착수.

## 1. 설계 원칙과 공식
- 실제 구조: Panda/Bow-tie 응력봉은 클래딩과 별개 physical group: stress_rod로 메시 생성. Bow-tie는 정확한 wedge가 아닌 타원 근사. 응력봉의 광학 굴절률은 기존 순수 SiO2로 명확하게 표시한 **미검증 proxy**이며 실제 B2O3 도핑 제품 근거가 아니다.
- 사용자가 반드시 공급할 재료값: T_use−T_stress_free = ΔT[K], 각 영역 E[Pa], ν, α[1/K], p11,p12 strain-optic 무차원, mechanical_source_url과 photoelastic_source_url. 기본값은 TODO/null, 추측한 수치는 사용 금지.
- PML 영역 제외, 선형 등방 탄성 free-surface boundary, generalized plane strain: εxx = ∂x ux, εyy = ∂y uy, εxy=(∂y ux+∂x uy)/2, εzz는 단면 일정한 미지수로 ∫σzz dA=0으로 결정. 외주 traction-free, rigid-body 3 DOF gauge만 고정.
- Lamé λL=Eν/[(1+ν)(1−2ν)], μ=E/[2(1+ν)]. σ = 2μ(ε−αΔT I)+λL(tr ε−3αΔT)I [Pa].
- 광탄성: ΔQii=p11 εii+p12(εjj+εkk), ΔQxy=(p11−p12) εxy, Q=ε_r^-1=I/n0²+ΔQ, ε_r=Q^-1. p11,p12는 무차원 strain-optic이며 stress-optic(m²/N) 상수와 다른 물리량.
- 삼각형별 복굴절 굴절률 텐서를 기존 HCurl(Nédélec)+H1 FEM에 anisotropic epsilon tensor로 연결. PML radial/tangential tensor transform 보존. 기존 STEP4/6a 코드 변경 최소화.
- core-material B는 코어 Q 면적평균에 따른 principal index 간격의 **재료 proxy**, 실제 편광모드 Re(neff_x)−Re(neff_y)가 아님. 복소 FEM에 quasi_x/y 둘이 식별되면 candidate B와 λ/B를 표시하지만 완전 모드 ID와 수렴 검증 전까지 NOT_CERTIFIED. 둘 중 하나라도 부족하면 TODO.
- 입력 SHA256은 디자인+재료 JSON+실제 메시+엔진 버전에서 산출, UTC 실행시간 포함. 디자인/재료 변경 시 웹 이전 결과 무효.

## 2. 구현 경로
| 소스 | 구현 |
| --- | --- |
| fiber_research_v1/stress_optic.py | Generalized plane strain sparse thermal FEM, axial force-free, material stress σ, strain ε, optical Q/ε tensor, core B proxy, explicit TODO |
| fiber_research_v1/geometry.py | Panda/Bow-tie 응력봉 별도 physical domain stress_rod |
| fiber_research_v1/presets.json, schema.py | 기존 25 설계 유지 + 응력봉 광학 surrogate를 별도 명시/검증 |
| fiber_research_v1/vector_fem.py | 입력된 anisotropic εxy/εxx/εyy/εzz triangle tensor를 Maxwell QEP/PML에 연결 |
| fiber_research_v1/pm_cli.py | design JSON + 출처 있는 material JSON → Gmsh → stress FEM → optical tensor → vector modes + output JSON/NPZ |
| web/research_step6b_v082.mjs | 재료 TODO 편집/저장, 온도/계수 근거 패널, 응력·국소 B 지도, material/modal B 분리, 결과 무효화 |
| tests/test_step6b_stress_optic.py | 12개 열팽창/기계/광탄성/입력 가드 테스트 |
| tests/test_step6b_vector_smoke.py | 3개 tensor anisotropy + 실제 Gmsh Panda stress + Maxwell QEP smoke |

## 3. 검증표 (공식 단위/스모크 vs 독립 수치검증)
| 체크 | 허용 및 예상값 | 판정 |
| --- | --- | --- |
| 균일 열팽창 | 동일 E,ν,α인 자유 단면: εxx=εyy=εzz=αΔT, σ≈0 | PASS (synthetic analytical) |
| 냉각 ΔT=0 | 잔류응력→0 | PASS |
| 이종 응력봉 열팽창 | rod α 차이 → 비영 기계 응력 | PASS (physical trend) |
| ∫σzz dA | axial free resultant≈0 | PASS |
| Q 광탄성 | ΔQxx−ΔQyy=(p11−p12)(εxx−εyy), Qxy=(p11−p12) εxy | PASS |
| 유효 상수 필수 | E,ν,α/출처 누락 시 reject; p 미제공 시 optical TODO | PASS |
| PML 구분 | 실제 physical domain만 기계 FEM, PML photoelastic δ ε=0 | PASS |
| 입력 무결성 | geometry/rod labels/materials/T값 + 실제 mesh SHA | PASS (기능 시험) |
| 전체 연결 | Gmsh Panda rod → thermo FEM → anisotropic Maxwell eigen candidate | PASS (SMOKE ONLY) |
| CI | Python tests + JS syntax + 25 presets schema | PASS, 58 tests (최종 run 참고) |
| 실제 PANDA optical/material | stress-rod composition-specific Sellmeier, α,E,ν,pij 출처 | EVIDENCE_REQUIRED |
| 실제 modal birefringence | neff_x, neff_y 독립 COMSOL/실측과 mode overlap 검증 | NOT PASS |
| PML / mesh | neff, B의 메시/경계 수렴 및 스퓨리어스 분리 | NOT PASS |
| PMD 실제 물리 검증 | DGD/statistical PMD vs stress-birefringence coupling | STEP6f 미착수 |

**주의:** PASS는 단위·synthetic 물리 예제·small FEM smoke에만 해당. PM 성능 정량·COMSOL 모드 고유값 검증을 통과했다는 의미가 아님.

## 4. 실행
설치: COMSOL/universal_silica_v01/fiber_research_v1/
- Python CLI: python cli.py presets --out presets
- 입력 안내: benchmarks/step6b_panda_materials_TEMPLATE.json 파일의 모든 TODO를 출처가 있는 실측/문헌값으로 채워 pm_materials.json 생성.
- 실제 계산: python pm_cli.py solve presets/polarization_maintaining__panda.json pm_materials.json --out out_pm --modes 4
- 기계 분석만: python pm_cli.py solve presets/polarization_maintaining__panda.json pm_materials.json --out out_pm_mech --mechanics-only

결과 파일: pm_stress_result.json, pm_tensor.npz, pm_mode_result.json (광학 계수 완비 시), pm_summary.json.
웹 STEP6b 패널에서 재료/요약/응력 파일을 각각 불러와 시각화한다. 브라우저 FEM 실행 없음.

## 5. 별도 한계 보고 / 다음 승인
[STEP6B_V082_LIMITATIONS_20261118.md](STEP6B_V082_LIMITATIONS_20261118.md) 참고. 다음 제안은 STEP 6c MCF N코어 supermode와 κ=(β_even−β_odd)/2 및 trench/pitch 최적화. 승인 전 착수 금지.

## 출처
- COMSOL Stress-Optical Effects, elastic → optical mode: https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.stress_optical/stress_optical.html
- COMSOL generalized plane strain: https://doc.comsol.com/6.4/doc/com.comsol.help.models.sme.stress_optical_generalized/stress_optical_generalized.html
- Isotropic p11,p12/p44 strain-optic relation: https://pmc.ncbi.nlm.nih.gov/articles/PMC12203310/
- PANDA thermal stress and FEM: https://www.sciencedirect.com/science/article/abs/pii/S1068520004000641
