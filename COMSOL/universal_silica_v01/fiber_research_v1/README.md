## v0.8.1 — STEP 6a 굽힘 해석 추가 (2026-11-18)

**현재 구현은 STEP6a만:** 공간적 등각 굴절률 n_eq=n(1+x_out/R_eff), 굽힘 방향별 반경-총손실 후보 곡선, PML 요구/동일 메시 복소장 overlap 추적, 선택적 문헌 근거 R_eff. HCF/PCF/MCF 굽힘 수치 정밀도 및 COMSOL V3 비교는 미인증입니다.

- [STEP6a 설계·검증표](../../validation_report/STEP6A_V081_DESIGN_VERIFICATION_20261118.md)
- [STEP6a 한계 보고서](../../validation_report/STEP6A_V081_LIMITATIONS_20261118.md)
- [v0.8.1 변경 이력](CHANGELOG.md)
- [웹 STEP6a 결과 탭](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#research-step6a-v081)

명령 예시(모든 기하/메시/PML 파라미터는 테스트 전용, 정량 인증 아님):

    python cli.py bend benchmarks/step6a_step_index_smoke_input.json --radii-mm 50 20 10 --directions-deg 0 90 --modes 2 --out out_bend

출력 out_bend/bend_sweep_result.json을 웹 STEP6a 패널에서 읽습니다. 기본 PEC 프리셋으로 손실을 계산할 수 없으며 반드시 독립 검증된 PML을 설정하세요. 응력 보정은 --reff-multiplier, --reff-source-url, --reff-source-note 세 옵션을 모두 지정할 때만 적용합니다. 원본 design JSON 변경 시 웹 결과가 무효화됩니다.

V3 공식 COMSOL Step-Index Fiber Bend는 axisymmetric와 cylindrical PML을 사용하므로 이 코드의 2D conformal 근사와 수치 동치가 확인되지 않았습니다. **STEP6b~6f, STEP7, STEP8 전체는 추가 승인 전 미착수입니다.**

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
