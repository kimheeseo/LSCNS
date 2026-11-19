## v0.8.7 / v0.8.6 — STEP 4·5 보완 (2026-11-19)

- STEP4 v0.8.6: 실제 6성분 물리 FE Gauss 적분(재료 라벨, PML 제외), Et Aeff/MFD2nd, Poynting 유리 비율, 출처 있는 quartic n2/gamma, 명시적 radial PII, D/S/λ0, 편광/HOMER·surface 후보 입력 보호. N2 Et⁴ 적분은 degree-8 quadrature를 사용한다. 불명확한 퇴화 편광을 개별 모드로 강제 연결하지 않고 부분공간 centroid로 추적한다.
- STEP5 v0.8.7: fixed-beta 벡터 H Maxwell PWE, triangular/honeycomb/square/ideal Kagome, Bloch 밴드와 light lines/표본 gap, cutoff/k convergence, finite triangular defect FEM handoff. 기존 scalar FDFD를 별도 선택으로 보존한다.
- [STEP4 검증·한계](../../validation_report/STEP4_V086_VERIFICATION_20261119.md) · [STEP5 검증·한계](../../validation_report/STEP5_V087_VERIFICATION_20261119.md) · [STEP4 설계](../../validation_report/STEP4_V086_DESIGN_20261119.md) · [STEP5 설계](../../validation_report/STEP5_V087_DESIGN_20261119.md)

```bash
python -m pip install -r requirements.txt
python step4_v086_cli.py result mode_result.json --context benchmarks/step4_v086/context_template.json --out metrics.json
python step4_v086_cli.py solve-sweep design.json --start 1510 --stop 1590 --points 9 --out sweep
python step5_v087_cli.py example --out step5_design.json
python step5_v087_cli.py analyze step5_design.json --out bands.json
python step5_v087_cli.py converge step5_design.json --out convergence.json
python step5_v087_cli.py handoff bands.json --candidate-index 0 --out finite_design.json
```

STEP4 result는 `python cli.py solve design.json --export-full-quadrature --out out`의 실제 적분 데이터가 필요하다. 희소 `field_samples`를 정밀 적분하지 않는다. 공기/유리 n2 기본값은 없으며 n2와 source가 모두 있어야 gamma가 계산된다. 비축대칭 vector Petermann-II·미확인 FM/HOM/표면 모드·극소 HCF 손실·complete bandgap은 자동 인증하지 않는다. 웹 상단 STEP4/STEP5에서 실제 계산 예시를 불러오며 브라우저 벡터 계산은 하지 않는다.

## v0.8.5 — STEP 3 풀벡터 모드 솔버 (2026-11-19)

STEP 1·2를 실제 웹 바로가기에 연결하고 기존 모듈 위에 혼합 벡터 FEM을 보완했습니다. 앱 v0.8.5 / 기본 scalar 엔진 v0.7.1 / 설계 schema_version=1.0.0을 구분합니다. 기존 preset 수치 및 기존 Colab/Yee 자산은 유지합니다.

- [설계](../../validation_report/STEP3_V085_DESIGN_20261119.md) · [알고리즘/출처](../../validation_report/STEP3_V085_ALGORITHM_20261119.md) · [검증표](../../validation_report/STEP3_V085_VERIFICATION_20261119.md) · [한계](../../validation_report/STEP3_V085_LIMITATIONS_20261119.md)
- Nédélec/H1 1·2차, 복소 β² sparse shift-invert + 기존 companion 교차 검사, PML H 복원, 6성분/잔차/전력, 독립 퇴화 모드 보존, FE 장 부분공간 추적, SBC target/선택적 반복을 추가했습니다.
- 웹 상단 **STEP 3 벡터 FEM**에서 Python의 실제 HE11/방사 예시를 적용하고 두 모드를 선택합니다. 입력 변경 시 이전 결과를 제거하며 불일치 결과는 거절합니다. 브라우저는 벡터 고유계를 계산하지 않습니다.
- V1/V4/V5는 **명시된 합성 원형 시험에서 PASS_SCOPED**입니다. PCF/HCF 극소 손실 또는 제품 dB/km 정확도 인증은 아닙니다. Ge/F 계수·C6v·정확한 일부 특수 CAD는 TODO입니다.

```bash
python -m pip install -r requirements.txt
OPENBLAS_NUM_THREADS=1 OMP_NUM_THREADS=1 python validate_step3.py --out out_step3
python cli.py solve benchmarks/step3_v085/guided_input.json --modes 2 --target-neff 1.468524 --out out_guided
python cli.py solve benchmarks/step3_v085/radiative_input.json --modes 2 --target-neff 0.981265 --out out_radiative
python cli.py study benchmarks/step3_v085/radiative_input.json --modes 2 --target-neff 0.981265 --out out_study
```

설계 JSON의 선택적 `solver`에는 `element_order: 1|2`, `backend: beta_squared|companion`, `max_dofs`, `tolerance`, `seed`, `sbc_iterations` 등을 지정합니다. 기본 혼합 DOF 한도 80,000(예시는 명시적으로 120,000), 최대 허용 250,000이며 물리 정확도나 메모리 성공을 보장하지 않습니다. Colab에서도 같은 Python 명령을 실행하고 `mode_result.json`을 웹으로 가져옵니다. 전체 설치/시험/한계는 위 문서를 확인하세요.

## v0.8.4 — STEP 1·2 보완 (2026-11-19)

기존 A–H **25개 설계 수치·schema_version=1.0.0·벡터 고유계 코드는 유지**하고 입력·기하·메시 검사를 보완했습니다. 기존 웹 scalar 표시는 v0.7.1, 이 독립 설계/메시 모듈은 v0.8.4입니다.

- [설계](../../validation_report/STEP1_2_V084_DESIGN_20261119.md) · [검증표](../../validation_report/STEP1_2_V084_VERIFICATION_20261119.md) · [한계](../../validation_report/STEP1_2_V084_LIMITATIONS_20261119.md)
- STEP 1: 잘못된 타입/필수 재료/계산창·hole·nested tube·MCF trench·PM 간섭 검사, 근사 형상 메타데이터, JSON Schema 상세 검증.
- STEP 2: y-major 타원의 OCC 회전 처리, 실제 재료 n(λ) 메시 목표, 곡률, 도메인·PML·경계 라벨, 원형 벽 법선 교차 검사, 3단계 기하 메시 실험.
- 웹 연구 프리셋 패널의 **STEP 1·2 · 설계·메시 검사 v0.8.4**에서 Python `mesh_preview.json`을 불러옵니다. 표시 요소는 최대 6,000개 표본이며 라벨 경계는 전체 표시합니다. 입력 수정 시 이전 그림을 무효화합니다.

    python cli.py presets --out presets
    python cli.py inspect presets/anti_resonant__nanf.json --out inspect_nanf
    python cli.py mesh presets/anti_resonant__nanf.json --out out_mesh
    python cli.py mesh presets/anti_resonant__nanf.json --resolution-policy wavelength --out out_mesh_strict
    python cli.py mesh-study benchmarks/step12_bragg_mesh_input.json --out out_mesh_study
    python validate_step12.py --out validation_step12.json

메시 정책:
- `boundary_refined`(기존 JSON 기본): 기존 벌크 메시 크기를 유지하며 재료별 파장 기준 미달을 보고합니다. 기본 25개 모두 벌크에서 이 엄격 파장 기준에 미달합니다.
- `wavelength`(명시적 선택): 실제 사용 재료 n(λ)의 가장 작은 λ/(n·요소수)를 전 도메인에 적용합니다. Gmsh 목표 길이 0.8 안전계수, 실측 최대 변 허용비 1.35, 원형 벽 표본 coverage 0.99는 공개된 수치 검사 정책이며 물리 상수가 아닙니다. 큰 메시를 만들 수 있고 기존 벡터 솔버 DOF 상한을 없애지 않습니다.

검증: Python 전체 83개 통과(기존 67개 포함), 25/25 기본 기하 생성·CAD 대비 면적·75개 원형 벽 표본 검사 통과. 이는 **기하/메시 검사**이며 neff·Im(neff)·손실·완전 밴드갭 인증이 아닙니다. 일부 명칭의 정확한 CAD와 C6v는 TODO입니다. 아래 v0.8.0–0.8.3 설명은 기존 이력으로 보존합니다.

## v0.8.3 — STEP 5 광자결정 클래딩 Bloch-Floquet (2026-11-18)

4종 단위셀의 **실제 2D 스칼라 Bloch 고유해석**과 밴드/표본 갭 후보·light line을 구현. 브라우저는 계산하지 않고 Python JSON 시각화만 수행한다. 기본 예시 JSON은 CI Python이 실제 계산했고 **갭 후보가 0개**로 나왔으며, 결과를 임의로 채우지 않는다.

- 설계·검증표: [STEP5](../../validation_report/STEP5_V083_DESIGN_VERIFICATION_20261118.md)
- 한계 보고서: [STEP5 LIMITATIONS](../../validation_report/STEP5_V083_LIMITATIONS_20261118.md)
- [GitHub Pages 광자결정 STEP5 설계 탭](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#research-step5-v083)
- [Python 실계산 예시 JSON](../web/research_step5_example_v083.json)

    python step5_cli.py example --out step5_design.json
    python step5_cli.py analyze step5_design.json --out out_step5/step5_bloch_result.json
    python step5_cli.py handoff out_step5/step5_bloch_result.json --candidate-index 0 --defect-cells 7 --out step5_finite_pbgf_design.json

handoff에는 계산된 진짜 sampled gap **후보**가 있어야 하며 없으면 생성할 수 없다. 삼각 격자만 기존 finite PBGF solver로 연결. full-vector periodic Maxwell, COMSOL V2/V4/V5/V6는 **미인증**. STEP6a, STEP6b 기존 구현은 변경하지 않았다.


## v0.8.2 — STEP 6b PM 열응력·광탄성 (2026-11-18)

**이번 단일 단계:** Panda/Bow-tie stress_rod 물리영역 분리, generalized plane strain 냉각 열응력, 사용자가 출처를 명시한 무차원 strain-optic p11/p12로 광학 굴절률 텐서 생성, anisotropic Maxwell mode candidate 연결. **PM 제품 수치 인증 아님.**

- [STEP6b 설계 및 검증표](../../validation_report/STEP6B_V082_DESIGN_VERIFICATION_20261118.md)
- [STEP6b 별도 한계 보고서](../../validation_report/STEP6B_V082_LIMITATIONS_20261118.md)
- [모든 변경 이력](CHANGELOG.md)
- [GitHub Pages STEP6b 열응력/복굴절 결과](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#research-step6b-v082)

재료값은 임의로 채우지 않습니다. benchmarks/step6b_panda_materials_TEMPLATE.json의 null 필드를 실제 PANDA/Bow-tie 제품의 조성/실측·문헌 근거 있는 값으로 채워 pm_materials.json을 만드세요. E[Pa],ν,α[1/K],p11,p12(무차원 strain-optic),source_url,ΔT[K] 입력이 필요합니다. Default stress_rod optical index는 **순수 SiO2 설명용 surrogate이며 미검증 실제 rod 조성**입니다.

    python cli.py presets --out presets
    python pm_cli.py solve presets/polarization_maintaining__panda.json pm_materials.json --out out_pm --modes 4
    python pm_cli.py solve presets/polarization_maintaining__panda.json pm_materials.json --out out_mechanics --mechanics-only

출력 파일: pm_stress_result.json, pm_tensor.npz, pm_mode_result.json(광탄성 모두 입력된 경우), pm_summary.json. 웹 결과 뷰어에서 PM 재료 JSON·요약·응력 지도 결과를 가져올 수 있습니다. Web은 Python FEM을 실행하지 않습니다.

기계 학술 공식과 synthetic small mesh smoke 58개 테스트 PASS, **제품 PM modal B의 독립 COMSOL/PML/mesh 수렴 미검증**. STEP6c~6f, STEP7/STEP8 전체 구현은 사용자 별도 승인 전 진행하지 않습니다.

## v0.8.1 — STEP 6a 굽힘 해석 추가 (2026-11-18)

**현재 구현은 STEP6a만:** 공간적 등각 굴절률 n_eq=n(1+x_out/R_eff), 굽힘 방향별 반경-총손실 후보 곡선, PML 요구/동일 메시 복소장 overlap 추적, 선택적 문헌 근거 R_eff. HCF/PCF/MCF 굽힘 수치 정밀도 및 COMSOL V3 비교는 미인증입니다.

- [STEP6a 설계·검증표](../../validation_report/STEP6A_V081_DESIGN_VERIFICATION_20261118.md)
- [STEP6a 한계 보고서](../../validation_report/STEP6A_V081_LIMITATIONS_20261118.md)
- [v0.8.1 변경 이력](CHANGELOG.md)
- [웹 STEP6a 결과 탭](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#research-step6a-v081)

명령 예시(모든 기하/메시/PML 파라미터는 테스트 전용, 정량 인증 아님):

    python cli.py bend benchmarks/step6a_step_index_smoke_input.json --radii-mm 50 20 10 --directions-deg 0 90 --modes 2 --out out_bend

출력 out_bend/bend_sweep_result.json을 웹 STEP6a 패널에서 읽습니다. 기본 PEC 프리셋으로 손실을 계산할 수 없으며 반드시 독립 검증된 PML을 설정하세요. 응력 보정은 --reff-multiplier, --reff-source-url, --reff-source-note 세 옵션을 모두 지정할 때만 적용합니다. 원본 design JSON 변경 시 웹 결과가 무효화됩니다.

V3 공식 COMSOL Step-Index Fiber Bend는 axisymmetric와 cylindrical PML을 사용하므로 이 코드의 2D conformal 근사와 수치 동치가 확인되지 않았습니다. **STEP6b는 v0.8.2에서 별도 수행되었으며 STEP6c~6f, STEP7, STEP8 전체는 추가 승인 전 미착수입니다.**

## v0.8.0 — STEP 4 ONLY (2026-11-18)

**현재 구현 범위:** signed dB/m 및 dB/km, Et-only Aeff/second-moment MFD, axisymmetric Petermann-II, overlap-gated D/S/λ0, 입력된 n₂ 기반 γ, HCF glass power/n₂ 가중 함수, HOMER, B/LB, ARROW 공진선, anticrossing 후보. 모든 공식·단위는 [physical_metrics.py](physical_metrics.py) 최상단에 고정했습니다.

- [STEP4 설계·검증표](../../validation_report/STEP4_V080_DESIGN_VERIFICATION_20261118.md)
- [STEP4 별도 한계 보고서](../../validation_report/STEP4_V080_LIMITATIONS_20261118.md)
- [버전별 변경 이력](CHANGELOG.md)
- [정적 웹 결과 확인](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#research-step4-v080)

STEP4 후처리 사용법:

    python cli.py solve presets/index_guiding__step.json --out out_step --modes 2
    python step4_cli.py result out_step/mode_result.json --out step4_metrics.json
    python step4_cli.py solve-sweep presets/index_guiding__step.json --start-nm 1500 --stop-nm 1600 --points 7 --out out_sweep --modes 2

두 번째 명령의 n₂는 **공식 자료/측정값을 명시적으로 제공하는 경우에만** --n2-m2-W로 지정하세요. 가스/유리 n₂가 없는 HCF γ는 TODO입니다. 모든 D/S/λ0 스윕은 동일 메시 복소장 overlap 검사를 거치며 인접 ID가 불명확하면 계산을 중단합니다.

**실제 물리적 정밀도는 아직 미검증:** GitHub Actions 수식/스모크 테스트 통과와 HE11·PML/mesh 수렴·HCF confinement loss 인증은 다릅니다. STEP5 및 다른 단계는 승인이 있기 전까지 착수하지 않았습니다.

# LSCNS Fiber Research v1 — STEP 1–3

**Prototype source code; not numerically certified.** The existing web scalar radial FEM, browser Yee-FDFD, web geometry and project schema 3 remain unchanged. New Python solver and design JSON are intentionally separate.

## Python 3.11 quick start

    cd COMSOL/universal_silica_v01/fiber_research_v1
    python -m pip install -r requirements.txt
    python -m unittest discover -s tests -v
    python cli.py presets --out presets
    python cli.py validate presets/index_guiding__step.json
    python cli.py mesh presets/index_guiding__step.json --out out_step
    python cli.py solve presets/index_guiding__step.json --out out_step --modes 2
    python cli.py study presets/anti_resonant__nanf.json --out out_nanf_study --modes 1

With package import from LSCNS root, use python -m COMSOL.universal_silica_v01.fiber_research_v1.cli. Source-file execution also works.

## Model definition and coverage

25 versioned design presets, eight fiber families A-H: index-guided 5, solid PCF 5, PBGF 2, Kagome 1, anti-resonant HCF 5, Bragg 1, MCF 3, polarization-maintaining 3. All parameters are engineering starting values, not qualified product designs. Save/load via CLI and schema.py; validation rejects invalid nesting and unsupported dopant Sellmeier claims.

Geometry uses gmsh OCC fragments, material physical groups, circular PML annulus, local thin-wall sizing, mesh diagnostics, and outputs geometry.msh, mesh.npz and mesh_report.json. Some variants are **illustrative approximations**: Kagome hypocycloids, conjoined and negative-curvature designs do not yet have exact multi-arc tube interfaces; graded/multilayer presets use piecewise regions; bow-tie stress rods are not thermoelastic calculations. Endlessly-single-mode diagrams and PBGF Bloch bandgap maps are **not implemented**.

Full vector is a 2D HCurl(Nedelec first-kind order 1) transverse E and H1(P1) axial Ez mixed quadratic-beta eigenproblem. It solves with SciPy sparse shift-invert of a companion linearization. Polynomial radial PML uses coordinate-transformed anisotropic eps/mu; SBC is a first-order Robin approximation at a fixed target neff, not exact outgoing DtN. Mode outputs: Re/Im(neff), total dB/km sign-preserved, six complex E/H components sampled at triangle centroids, relative QEP residual, weak Gauss residual, polarization fraction and PML confinement ratio. This **does not** prove all spurious modes have been removed.

Pure silica Sellmeier: Malitson 1965 https://doi.org/10.1364/AO.4.000901. Ambient air n=1.00027 is an uncalibrated assumption, constants are user-specified. Ge/F doped Sellmeier requires a composition-dependent authoritative coefficient table: TODO and explicit error. PM stress-optic anisotropy and bandgap are not solved.

## Acceptance and safety

exp(i beta z - i omega t), ntilde=n+i*kappa. Im(neff)>0 means absorption. Total loss dB/km = (20/ln 10)*2pi/lambda_m * Im(neff)*1000. Negative loss is NOT silently made positive.

A converged linear algebra residual is necessary, not sufficient. For any published HCF/PCF confinement loss, run at least 3 meshes, multiple physical windows, 3 PML thicknesses, 3 strengths and 3 polynomial orders, match modes using field overlap, compare HE11/independent reference, examine boundary energy and Gauss residual. The study command runs numerical sweeps but uses nearest-index matching only; even a sensitivity PASS remains independent-benchmark-unverified. Do not claim COMSOL parity or measured product specifications.

Environment note: gmsh/scikit-fem runtime was not available during source authoring in this conversation. New FEM code is NOT claimed to have executed or passed its physical validation. Run in Colab/local with dependencies and inspect output.

API docs: https://scikit-fem.readthedocs.io/en/stable/api.html and https://gmsh.info/doc/texinfo/gmsh.html

