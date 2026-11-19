# STEP 5 v0.8.7 — 검증표·한계 (2026-11-19)

| 시험 | 공개 기준 | 실제 결과 |
| --- | --- | --- |
| 4종 lattice 균질 매질, beta=0/0.4/1.2 두 편광 exact folded spectrum | 최대 오차 ≤1e-10 | 2.22045e-16 |
| Maxwell normwise 잔차 / divergence | 각각 ≤1e-10 | 9.53078e-19 / 2.04814e-18 |
| beta=0 square, 독립 real-space TE/TM finite difference N=96 | 저차 6밴드 최대 상대차 ≤2% | 0.194504% |
| Bloch boundary phase, Hermitian, Fourier 면적분, nested variational cutoff | 위상·Hermitian 1e-10, 면적분 5e-4, nested 고유값 비증가 | 새 7개 단위 시험 통과 |
| 실제 계산한 예시의 표본 간극 | 없으면 0개 표시, 있으면 후보만 표시 | 1개 후보 |
| cutoff 3/4/5 및 k-mesh 3/5/7 | 실제 간극 edge와 표본 변화 공개 | 6개 경우 계산 |
| 실제 간극 → 결함 CAD → 전체 단면 FEM | 지정 beta/k0를 target_neff로 전달, 잔차·유한 장 출력 | finite_FEM_example 원시 JSON 참조; candidate 상태, 손실 미인증 |

예시: 삼각 원형 air hole d/Λ=0.9, n_host=1.45/n_hole=1.00, Λ=1.5µm, beta_norm=0.6/1/1.4, 12밴드, cutoff=5. 명시적 상수 index 연구 시험이며 실제 실리카 Sellmeier/상용 HCF 제품을 의미하지 않는다. βnorm=1.4의 4/5밴드 표본 후보가 생성됐다. `convergence.json`에서 cutoff 증가에 따른 edge 이동과 k sampling 영향을 확인해야 한다. 기본 예시의 간극을 complete gap으로 인증하지 않는다.

## 재현·툴 사용

```bash
python step5_v087_cli.py example --out step5_design.json
python step5_v087_cli.py analyze step5_design.json --out bands.json
python step5_v087_cli.py converge step5_design.json --out convergence.json
python step5_v087_cli.py handoff bands.json --candidate-index 0 --rings 4 --out finite_design.json
python cli.py solve finite_design.json --target-neff <생성된 target_neff> --out out_modes
```

`--solve-fem`으로 handoff 뒤 FEM을 실행할 수도 있다. 기본 handoff의 thin-wall 해상도와 큰 설정은 DOF 한도를 넘을 수 있으며 한도를 자동 해제하지 않는다. computational radius는 명시값 또는 (rings+3)*pitch를 사용하며 `--domain-radius-um`/JSON으로 제어한다. 이는 물리 상수가 아닌 공개된 계산창 정책이며 독립 window/PML convergence가 필요하다.

실제 FEM 예시는 2-ring, core surround=0µm, 4 elements/wavelength, wall target=3, N1/P1의 명시적 저해상도 수치 사례다. 원래 프리셋 수치를 바꾸지 않았다. 표준 0.15µm surround/기존 정밀 mesh의 기본 80,000 DOF 초과를 숨기거나 성공으로 보고하지 않는다. FEM 예시 4개 모드는 candidate_unverified이고, 코어 disk Poynting 비율은 약 1.76–7.70%로 낮아 이 예시에서 core-confined mode를 확인하지 못했다. 이 결과를 코어 구속 성공으로 표시하지 않는다. core disk Poynting fraction과 실제 neff·잔차는 validation.json에 기록했다. 높은 core fraction만으로 손실/모드 정체성/밴드갭 confinement를 인증하지 않는다.

웹 STEP5의 단위셀/밴드/결함 탭에서 vector PWE와 legacy scalar FDFD를 구분해 선택한다. 새 JSON Schema는 `step5_vector.schema.json`, schema_version=1.0.0이다. 해석 결과 import/예시 적용, 후보 선택, 결함 JSON 저장 및 공통 연구 편집기 적용, 사전 계산 FEM/STEP4 물리량 import를 지원한다. 모든 벡터/밴드 고유해석은 Python에서 수행한다.

## 한계·미확인

- dense PWE cutoff≤8, 실제 circle analytic Fourier; 실수·등방성·비분산 ε, µ=1만 지원한다. Ge/F Sellmeier·복소 흡수·자기매질·일반 tensor dispersion의 ω 비선형 고유계는 구현하지 않았다.
- MPB 이론을 근거로 작성한 독립 NumPy/SciPy 구현이다. MPB 실행 비교는 NOT_RUN이며 COMSOL 재현이나 MPB 동등성으로 표시하지 않는다.
- 실공간 TE/TM 비교는 beta=0 square에서만 독립 수치 교차 검증이다. 비제로 beta의 모든 실험 PCF를 검증한 것이 아니다.
- cutoff/k mesh finite extrema는 full Brillouin zone의 수학적 gap proof가 아니다. 고차 reciprocal 파 및 날카로운 고대비 경계에서 느린 수렴이 가능하다. 수렴 데이터가 있어도 complete-gap 인증 flag는 false다.
- honeycomb·Kagome은 다중 원형 site motif이고 실제 strut/hypocycloid HCF의 CAD가 아니다. finite CAD handoff는 triangular 원형 공기 구멍 7/19-cell 근사만 지원한다. 19-cell은 최소 3 rings가 있어야 남는 클래딩 hole이 있다.
- beta/k0 gap 중심 target은 검색 시작점이다. 유한 defect mode의 β가 정확히 일치하거나 해당 core mode가 존재한다는 보증이 아니다. 실제 core localization·모드 식별·mesh/window/PML 수렴과 광섬유별 문헌 비교가 뒤따라야 한다.
- 기존 scalar FDFD와 v0.8.0/0.8.3 파일은 보존한다. 새 vector 결과를 scalar 결과로 또는 그 반대로 표시하지 않는다.

최종 CI·브라우저 증거는 STEP4/5 통합 완료 보고서에 추가한다.
