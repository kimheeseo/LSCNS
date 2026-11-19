# STEP 4 v0.8.6 — 검증표 및 한계 (2026-11-19)

| 시험 | 기준 | 실제 결과 |
| --- | --- | --- |
| Gaussian Et 면적·second moment·radial Petermann-II | 해석 Aeff=πw², MFD=2w, 수치 오차 <0.5% | 새 12개 단위 시험 통과; w=4µm synthetic Gaussian |
| 손실 단위·부호·흡수 구분 | dB/km=1000 dB/m, 복소 neff 부호 보존, 흡수 포함이면 radiation-only null | 통과 |
| 실제 물리 FE 적분 | PML label 제외, 6성분·가중치·재료 라벨, Poynting partition 합=1 | 실제 방사 원형 모드 2개에서 통과 |
| 모드 추적 | 순서 permutation 및 퇴화 기저 회전 불변, 개수/격자 변경 차단 | 통과; 퇴화군은 centroid만 보고 |
| 실제 PEC 원형 TE11 D | 정확한 complex neff(λ), max |ΔD|≤0.1 ps/(nm km) | 0.00642798125 |
| 실제 PEC 원형 TE11 S | max |ΔS|≤0.01 ps/(nm² km) | 4.31391979e-06 |
| 파장 간격 10/20nm | max |Dfine−Dcoarse|≤0.1 | 2.66530684e-07 ps/(nm km) |
| 인접 부분공간 overlap | ≥0.75 | 최소 0.999999999999998 |

원시 수치: `fiber_research_v1/benchmarks/step4_v086/validation.json`, 실제 PEC FEM `example_metrics.json`/`example_sweep.json`, 실제 방사 해의 유리/공기 partition `radiative_metrics.json`. 실제 섬유 실험 데이터가 아닌 명시적 합성 원형 Maxwell 시험이다.

## 구현·사용

`python step4_v086_cli.py result mode_result.json --context context.json --out metrics.json`은 실제 6성분 Gauss quadrature를 요구한다. `python cli.py solve design.json --export-full-quadrature --out out`으로 만든다. `python step4_v086_cli.py solve-sweep design.json --start 1510 --stop 1590 --points 9 --target-neff 1.442 --out out_sweep`는 공통 FE probe와 단일 파장 메시로 실제 FEM을 계산한다. 일반 파라미터는 `--parameter parameters.<key>`로 별도 재메시 후 같은 물리 probe에서 추적한다. λ 외 파라미터에는 D/S를 만들지 않는다.

`context.json`의 `nonlinear_materials`는 각 실제 material label에 `{n2_m2_W: 값, source: 원문 URL 또는 명시된 실험 조건}`을 제공해야 한다. 기본 n2는 없고 전부 누락 시 TODO이다. `material_roles`는 라벨을 air/glass로 명시한다. `polarization_pair`와 `polarization_identity_source`, FM/HOM `mode_identities`/`identities_source`/`losses_converged`는 독립 확인 자료가 있어야 공급한다. `surface_pair`/`surface_identity_source`를 제공하면 추적 branch의 실제 유리 Poynting fraction 교환과 neff gap 최소로 후보를 선별한다. 이는 서명된 인증 입력이 아니라 사용자가 명시적으로 공급하는 provenance다.

## 한계

- Im(neff)는 물질 흡수와 방사 성분의 합일 수 있다. 보존된 STEP3 수치 검증을 모든 임의 HCF 손실 인증으로 전이하지 않는다.
- HCF n2_eff는 |Et|⁴ 가중 scalar Kerr 근사다. 단순 파워 비율로 n2를 가중하지 않는다. gas 조성·압력·온도·파장 및 실리카 n2 출처는 사용자 자료가 필요하다.
- 일반 2D 벡터 모드에 standardized Petermann-II를 임의 할당하지 않는다. 직접 제공한 축대칭 scalar radial profile에는 near-field derivative 식을 계산하며 finite-window/derivative convergence가 필요하다. second-moment와 별도 이름으로 표시한다.
- HOMER와 surface anticrossing 알고리즘은 있어도 실제 FM/HOM/표면 모드 식별 및 독립 loss convergence가 없으면 null/TODO이다. 자동 손실 인증이나 실제 HCF anticrossing 확인을 주장하지 않는다.
- overlap 임계값 0.75 및 degeneracy tolerance 1e-6은 수치 정책이다. topology/모드 수 변화는 D를 차단하며, 편광 기저가 불명확한 퇴화군의 개별 모드 D/B를 임의 계산하지 않는다.
- ARROW는 dispersive Sellmeier를 넣어 λ=2t√(ng²−na²)/m의 근을 자기일치 계산해도 얇은 벽 근사다. 실제 FEM 손실 피크와 일치를 자동 인증하지 않는다.

STEP4의 로컬 검증 이후 사용자가 함께 승인한 STEP5를 진행했다. 최종 CI·브라우저 기록은 STEP4/5 통합 완료 보고서에 추가한다.
