# LS Wave Optics v0.8.8 — STEP6 확장 해석

2026-11-20 (KST). 기존 v0.7.1 scalar, v0.8.6 mixed vector, STEP1–5, 이전 STEP6a/b 코드와 CLI를 보존한다. 새 실행기는 `step6_cli.py`이며 계산은 로컬/Colab Python, 웹은 입력 JSON 저장·결과 가져오기·표/차트/Pareto 표시만 수행한다.

## 설계 및 구현 맵

| 구분 | 추가 구현 | 검증·적용 범위 |
|---|---|---|
| 6a | `step6_analysis_v088.bend_sweep`: 동일 메시 PML, 등각 굴절률, R·방향 Cartesian sweep, 공통 실제 FE probe와 principal angle로 1모드/2차원 퇴화 부분공간 연결 | 소곡률 2D 근사. 총손실과 직선 대비 총손실 차이를 보고하며 독립 굽힘손실 인증 아님. 방향별 평균/범위 및 1회전 dB 저장 |
| 6b | 기존 generalized plane strain 열탄성 FEM과 p11/p12 기반 epsilon 텐서 → 벡터 FEM → 공간 intensity overlap과 직교 편광 후보 검사 | E, ν, α, ΔT, p11/p12 및 출처 필수. 실제 PANDA 조성 상수는 미확인 TODO. 예시는 합성 값 |
| 6c | N코어 실제 supermode, 실제 E/H 적분 core power, core-center 전계, unique principal-Jones overlap + even/odd envelope κ | 64모드까지 요청 가능하나 DOF/메모리 제한. 2N개 반환이 완전한 guided spectrum 증명은 아님. κ는 원점 대칭 동일 2코어 무굽힘에 한정 |
| 6d | 동일 파장 실제 FEM grid, Et overlap, 축 offset, 소각 phase tilt, padded outgoing angular spectrum gap, 선택적 scalar Fresnel | Gaussian 독립해 검증. SMF↔HCF 실제 FEM 후보 예시 포함. full E/H 경계 matching·융착 collapse·산란/다중 반사 미구현 |
| 6e | 다차원 유한 grid, 명시적 min/max·단위·제약, 실패 점 제외, Pareto, 목적별 최적 입력 저장, checkpoint/resume | 각 점 실제 FEM. 고유값 번호로 정한 목적은 모드 혼동 가능. guided 물성 목적에는 추가 overlap branch 추적 필요. Bayes/GA 미구현(선택 요구) |
| 6f | 기존 PMD 시연용 표시 유지. 부호 있는 Δn(λ) 추적 결과 → Δng → 균일 무결합 DGD | 통계 PMD[ps/√km] 계산과 구별. longitudinal random Jones/correlation 경로를 별도 제안 |

`step6_fields_v088.py`: 실제 FE probe 디코딩, 부분공간·공간 overlap, regular Cartesian grid export. `vector_solver_v085.py` 요청 모드 상한만 8→64로 확장했으며 고유계·진단·기존 호출 기본값은 유지했다. mixed FEM 엔진은 0.8.6, 확장 앱은 0.8.8로 구분한다. STEP6 입력은 `schema=LSCNS_step6_input_v088`, `schema_version=1.0.0`, `operation`을 사용한다. 각 설계는 기존 versioned design JSON을 그대로 포함한다.

## 공식과 단위

공식은 각 Python 모듈 첫 docstring에 고정했다.

- α[dB/m]=(20/ln10)(2π/λ_m)Im(neff). 수치 오차로 음수인 값을 절댓값으로 숨기지 않는다.
- x_out=x cosθ+y sinθ [µm], n_eq=n(1+x_out/(1000 R_eff_mm)). θ는 +x에서 반시계방향, 바깥쪽 방향. R_eff=f_source R. 경험 계수 기본값 없음. 사용 시 multiplier, 문헌 URL, 섬유/파장/조건 적용 설명을 모두 입력한다. 2D conformal과 COMSOL 축대칭 bend는 다른 정식화다.
- PM 광탄성: ΔQ_ii=p11 ε_ii+p12(ε_jj+ε_kk), ΔQ_xy=(p11−p12)ε_xy, εr=Q⁻¹. B=|Re(nx)−Re(ny)|, L_B=λ_m/B. B=0이면 beat length는 JSON null(무한대 한계)이다.
- β=k0 neff [m⁻¹], κ=(β_even−β_odd)/2 [m⁻¹]. 동일·대칭·무손실 근사에서 Lc=π/(2|Reκ|) [m]. 일반 N코어·detuning 상태에서 임의 인접 neff 차이를 κ로 부르지 않는다. ring의 pitch는 반지름이며 실제 최근접 코어 간격을 별도 저장한다.
- η_Et=|∫Et1·Et2* dA|²/(∫|Et1|² dA∫|Et2|² dA), IL=−10log10η. MFD=2w Gaussian 크기 불일치 η=(2w1w2/(w1²+w2²))². 두 실제 필드에 MFD 불일치를 다시 곱하지 않는다.
- 각도: exp(ik_gap[x sinθx+y sinθy]), |θ|≤10°. 갭: exp(iz sqrt(k²−kx²−ky²)), outgoing 복소 root, FFT pad 2–4. 유한 창/샘플링/패딩 수렴 필수.
- Fresnel T=4n1n2/(n1+n2)². 사용자가 제공한 실수 계면 n에만 적용한다. butt: [n1,n2], gap: [n1,n_gap,n2]. neff를 자동 계면 굴절률로 사용하지 않는다. 미입력 시 Fresnel 및 합산 손실 null.
- Δng=Δn−λ dΔn/dλ, DGD_ps=L_m|Δng|/299792458×10¹². 절댓값 B를 미분하면 축 교차에서 잘못된 결과가 나므로 signed nx−ny 및 추적 근거를 요구한다.

## 실행

```bash
python -m pip install -r requirements.txt
python validate_step6_v088.py --out benchmarks/step6_v088
python step6_cli.py bend benchmarks/step6_v088/bend_input.json --out bend_result.json
python step6_cli.py pm benchmarks/step6_v088/pm_input.json --out pm_result.json
python step6_cli.py mcf benchmarks/step6_v088/mcf_input.json --out mcf_result.json
python step6_cli.py sweep benchmarks/step6_v088/grid_input.json --out grid_result.json
python step6_cli.py field-grid benchmarks/step6_v088/field_grid_input.json --out field_grid.json
python step6_cli.py splice benchmarks/step6_v088/splice_input.json --out splice_result.json
python step6_cli.py pcf-reference --out pcf_reference.json
```

웹 STEP6 버튼에서 현재 STEP1/2 설계를 포함한 실행 JSON을 저장할 수 있다. PM에서는 먼저 입력 템플릿을 눌러 TODO를 실제 출처 값으로 채운다. `benchmarks/step6b_panda_materials_TEMPLATE.json`은 생산용 상수 기본값을 제공하지 않는다. 예시 불러오기는 예시 원 설계를 편집기에 적용한다. 별도 결과 파일은 현재 설계와 일치해야 하며 설계/입력 변경 시 이전 결과를 제거한다.

MCF/PM CLI는 요약 JSON과 `_modes.json`을 각각 저장한다. 실제 필드 grid는 `field-grid`에서 정한 같은 창/파장으로 두 설계를 각각 계산해 splice 입력의 field1/field2에 포함한다. 자동으로 단면/파장을 맞춰 만들지 않는다. 코드 함수/CLI의 target_neff와 요청 모드 수를 충분히 조정해야 한다. 기본 80,000 mixed DOF 자원 제한은 설계 solver 설정으로 명시적으로 조정하며 메모리 성공은 보장하지 않는다.

스윕 예시는 3코어 trench의 pitch×wall 2×2 그리드이고 mean core power와 최대 후보 총손실을 동시에 비교한다. 목적값은 각각의 반환 모드 집합에 대한 값이다. 실제 FM-only 또는 κ 최적화를 하려면 사용자가 모드 식별/편광 및 파장·형상 추적을 검증해야 한다. checkpoint의 전체 입력/evaluator 버전 hash가 다르면 재사용하지 않는다.

## 검증과 한계

실제 결과는 `benchmarks/step6_v088/validation.json`과 `COMSOL/validation_report/STEP6_V088_VERIFICATION_20261120.md`에 기록한다. 독립 Gaussian 크기·offset·각도·갭, Fresnel, 부호 Δn group-delay 항등식, 회전 퇴화 부분공간, 실패/제약/재개/Pareto 단위 시험을 실행한다. 열응력·MCF 합성 예시의 숫자는 상용 광섬유 성능을 뜻하지 않는다.

PM은 선형 등방 열탄성·일정 냉각·일반화 평면변형 가정이다. 점탄성 완화/인발 이력/3D warping 없음. total strain 광탄성 처리에는 free thermal strain도 포함하므로, 기준 n(T)와 일치하는 dn/dT 공통항 분리·검증은 TODO다. full xz/yz tensor도 지원하지 않는다. PM rod 상수/잔류응력 검증 전 정량 생산 예측에 사용하지 않는다.

PML/메시/창 수렴 전에는 음수·미세 Im(neff), bend difference, PM B 또는 극소 HCF 손실을 인증하지 않는다. 같은 FEM 엔진의 반복 실행은 COMSOL 독립 비교가 아니다. API의 64모드 한도는 19코어 전체 guided spectrum 정확도 보증이 아니다. twist 옵션은 helicoidal Maxwell 또는 z-dependent coupled-mode/Jones 해석이 필요하므로 아직 수치 지원하지 않는다.

PCF V2는 공식 예제의 정확한 `microstructured_optical_fiber_parameters.txt`가 공개 HTML 본문에 없어 미확인 입력을 TODO로 둔다. V6는 Folkenberg Table I 및 cutoff 식의 입력을 확인했으나 FM/HOM 수렴 검증을 통과했다고 표시하지 않는다. 문헌 cutoff fit을 FEM 결과로 대체하지 않는다.

## 근거 자료 (2026-11-20 확인)

| 용도 | 제목 / URL | 적용 범위 |
|---|---|---|
| 굽힘 | [Step-Index Fiber Bend, COMSOL 6.4](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.step_index_fiber_bend/step_index_fiber_bend.html) · 기존 bending.py conformal 문헌 URL | COMSOL 축대칭과 본 2D 근사 구별. R_eff 계수 특정 값은 미확인 |
| 응력광학 | [Stress-Optical Effects in a Waveguide, COMSOL 6.4](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.stress_optical/stress_optical.html) | 열탄성·응력광학 연결 근거. 해당 Si/SiO2 planar 예제 상수를 PANDA boron rod로 전용하지 않음 |
| supermode | [Directional Coupler, COMSOL 6.4](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.directional_coupler/directional_coupler.html) | symmetric/antisymmetric β 차이와 beating. GaAs square 예제이며 원형 MCF 수치 reference 아님 |
| 접속 | [Single Mode Fiber-to-Fiber Coupling, COMSOL 6.4](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.single_mode_fiber_coupling/single_mode_fiber_coupling.html) · [Understanding the mode overlap calculation, Ansys](https://optics.ansys.com/hc/en-us/articles/360034396834-Understanding-the-mode-overlap-calculation) | E/H power matching와 단순 Et 근사를 구별 |
| 갭 | [Matsushima & Shimobaba, Band-Limited Angular Spectrum Method, 2009](https://doi.org/10.1364/OE.17.019662) · [저자 기관 원문](https://kansai-u.repo.nii.ac.jp/records/11127) | outgoing angular spectrum 및 sampling 한계. 구현은 padded 기본 ASM이며 논문 bandlimit 알고리즘 전체 재현 아님 |
| PMD | [ITU-T G.650.2 (2015), Definitions and test methods for statistical and non-linear related attributes](https://www.itu.int/rec/T-REC-G.650.2-201508-I/en) | DGD·통계 PMD 정의. Δn만으로 random coupling coefficient 결정 불가 |
| PCF V2 | [Leaky Modes in a Microstructured Optical Fiber, COMSOL 6.4](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html) | neff 표는 확인, 외부 parameter txt의 정확한 숫자는 미확인 |
| PCF V6 | [Folkenberg et al., Experimental investigation of cut-off phenomena in non-linear photonic crystal fibers, 2003](https://arxiv.org/pdf/physics/0307009) | Table I: 7개 pitch·d/Λ, Eq1: α=2.80±.12, γ=.89±.02, threshold .406. FEM pass 아님 |

## 다음 제안

① V2 원본 parameters.txt 또는 검증된 COMSOL export 확보 → 입력 일치 외부 비교. ② V6 FM/HOM 및 cladding FSM cutoff 검증. ③ PM rod의 실제 조성·온도별 재료 상수와 측정 beat length 확인. ④ SMF↔HCF full E/H mode-matching 및 실제 splice 형상 해석. ⑤ longitudinal stress/twist correlation 데이터→Jones matrix cascading→DGD/PMD 분포 및 측정 비교. 다음 단계 승인은 별도로 받는다.
