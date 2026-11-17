# STEP 5 — 광자결정 클래딩 Bloch–Floquet | v0.8.3 | 2026-11-18

**구현: 스칼라 수치해석 프로토타입. 기능·합성 해석해 검증 PASS. 실물 PCF/COMSOL 및 full-vector 검증 NOT PASS.**

## 1. 요구와 설계

흐름: primitive 단위셀 → 고정 축방향 β의 ω–k_B 밴드 → β/k0 대 규격화 주파수 f의 갭 후보 → 결함 코어 파라미터 → 기존 full-section FEM design JSON.

- triangular: 60° primitive vectors, 공기 구멍 1개/셀.
- hexagonal: 60° Bravais lattice, 2-site honeycomb basis. 삼각 격자 1-site와 구별.
- square: 90° 정사각 격자, 1-site.
- kagome: 60° Bravais lattice, 3-site **원형 홀 근사**. 실제 strut/hypocycloid HCF가 아님.

입력: Λ[µm], d/Λ, n_air, n_host (실수·파장 무관), N×N, 밴드 수, β_norm, Bloch path, reciprocal k-mesh. 인접 구멍 비중첩 검사.

## 2. 수식, 물리·단위 계약

Convention E(x,y,z,t)=E(x,y) exp(iβz−iωt).

    −∇ₜ²E + β² E = k₀² n(x,y)² E
    k₀ = ω/c
    E(q+eᵢ) = exp(i k_B·aᵢ) E(q)

aᵢ는 실제 primitive vector, q는 fractional unit-cell coordinates. 복소 Bloch 위상 및 oblique metric g=A⁻¹A⁻ᵀ의 mixed derivative를 포함하는 9-point central finite difference. SciPy의 Hermitian generalized sparse eigenproblem K E = k₀² M E 고유해석. **2D scalar Ez surrogate이지 full-vector Maxwell FEM은 아니다.**

    f = Λ/λ = ωΛ/(2πc)
    b = βΛ/(2π)
    β/k₀ = b/f

고대칭점: 60° Γ–M–K–Γ; 90° Γ–X–M–Γ. reciprocal primitive parallelogram 전역 샘플을 추가한다. min_k f_(band+1) > max_k f_band는 **SAMPLED_GAP_CANDIDATE**로만 출력. 제한된 k-grid의 빠진 교차점, 수치 메시 오차 때문에 완전 밴드갭 증명 불가. Air/host light line은 β/k₀ = n_air / n_host.

## 3. 구현 파일과 실제 연결

| 구분 | 경로 | 역할 |
|---|---|---|
| 수치 | fiber_research_v1/step5_bloch.py | 격자·motif, sparse Bloch Hermitian operator, bands, gap candidates |
| CLI | fiber_research_v1/step5_cli.py | example / analyze / handoff |
| 테스트 | fiber_research_v1/tests/test_step5_bloch.py | 해석해·Hermitian·격자·handoff |
| Web | web/research_step5_v083.mjs, .css, _panel.html 명명 제외(실제 파일 research_step5_v083_panel.html) | 단위셀 설계/밴드 그래프/결함설계의 3개 탭 |
| 최종 모드 | 기존 fiber_research_v1/cli.py solve | 기존 삼각 pbgf 설계 JSON 입력 후 FEM 별도 실행 |

예시 CLI:

    cd COMSOL/universal_silica_v01/fiber_research_v1
    python -m pip install -r requirements.txt
    python step5_cli.py example --out step5_design.json
    python step5_cli.py analyze step5_design.json --out out_step5/step5_bloch_result.json
    python step5_cli.py handoff out_step5/step5_bloch_result.json --candidate-index 0 --defect-cells 7 --rings 4 --out step5_finite_pbgf_design.json
    python cli.py solve step5_finite_pbgf_design.json --out out_step5/core_modes --modes 4

마지막 두 명령은 실제 생성된 유효한 0번 갭 후보가 있어야 실행 가능하다. 예제의 후보가 없는 경우에는 건너뛴다. Handoff는 triangular만 허용하고 입력 조건·wavelength range가 맞아야 한다.

## 4. 검증표

| 항목 | 허용 조건 | 상태 |
|---|---|---|
| 60°/90° primitive cell area | √3Λ²/2 / Λ² | PASS (synthetic analytic unit) |
| motif nearest distance | tri Λ / honeycomb Λ/√3 / Kagome Λ/2 / square Λ | PASS (unit) |
| Bloch conjugate Hermiticity | max abs(K−K†) < 1e−10 | PASS (unit) |
| Homogeneous Γ eigenfrequency | f=b/n, abs error < 2e−6 | PASS (unit) |
| Homogeneous nonzero Bloch k | plane wave abs error < 6e−4 at N=22 | PASS (unit) |
| Gap resolution convergence | at least 3 N and 3 reciprocal k-mesh densities | NOT RUN |
| Full vector fixed-beta Floquet | independent Maxwell reference | EVIDENCE REQUIRED |
| V2 COMSOL leaky modes | neff / Im(neff) same mode, geometry and wavelength | EVIDENCE REQUIRED |
| V4/V5 PML + mesh | independent convergence table | NOT PASS |
| V6 PCF endlessly single-mode | peer-reviewed comparative mode cutoff | EVIDENCE REQUIRED |
| 4-lattice finite-core design | triangular 7/19 handoff only | PARTIAL |
| Existing Python test suite | 67 tests in CI at implementation checkpoint | PASS (software) |

**STEP 7의 PASS/FAIL 검증 탭 전체는 구현 완료로 표시하지 않는다.** STEP 6a/6b는 기존 별도 커밋이며 이 STEP5 작업에서 변경하지 않았다. 사용자가 STEP 8의 나머지 상세 사항을 전달하기 전까지 STEP 8은 추가로 진행하지 않는다.

## 공식 참고

- COMSOL Photonic Crystal band-gap: https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.bandgap_photonic_crystal/bandgap_photonic_crystal.html
- COMSOL Modeling Periodic Structures: https://doc.comsol.com/6.4/doc/com.comsol.help.woptics/woptics_ug_modeling.5.10.html
- COMSOL Leaky Modes in a Microstructured Optical Fiber: https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html
