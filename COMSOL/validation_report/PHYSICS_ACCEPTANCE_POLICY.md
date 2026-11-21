# 물리 검증 수용 정책 — 2026-11-21

소프트웨어 회귀 성공과 물리 gate를 분리한다. 전체 V1–V9가 PASS여야 all_required_passed=true이다. V8/V9는 공식/파서 범위이며 제품 인증이 아니다. 과거 V1/V4/V5의 합성 형상 PASS를 PCF/HCF에 전이하지 않는다.

## 원자료와 재현성

`physics_20261121/external_manifest.json`은 확보 여부를 나타내는 manifest다. 결측 파일과 허용오차는 null로 유지한다. 파일 사양은 `{path, sha256}`이고 path는 evidence root 내부 상대 경로다. 원본 txt/CSV 및 실제 장 파일을 사용하고 그림 추정치·반올림 straight neff로 대체하지 않는다. SHA256는 파일 동일성만 검증하며 외부 자료의 진실성이나 과학적 판단을 인증하지 않는다.

V2/V3 공통 필수 자료: reference_csv, calculated_csv, identity, convergence, model_manifest. V2는 exact_parameters 추가. V3는 raw_bent_csv, normalization_manifest, formulation_review 추가. model_manifest에는 COMSOL 버전, source_url, input_equivalence_reviewed_by가 필요하다. V2는 10개 모드를 모두 비교한다. 각 CSV의 부호 변환을 별도로 명시하고, Re_ppm/Im_relative/Im_absolute 허용값을 선언한다. 0/음수 Im를 절댓값으로 양수화하지 않는다.

## 모드 식별 및 독립 수렴 계약

identity JSON: reviewed_by, basis, mode_mapping 목록. 각 항목은 wavelength_nm, reference_mode_id, calculated_mode_id, reference_field 및 calculated_field의 파일 사양, minimum_subspace_correlation을 포함한다. 두 field 파일은 실제 좌표/가중치/복소 E를 tracking_fields 형식으로 가진 mode JSON 또는 modes 목록이다. gate가 실제 부분공간 correlation을 다시 계산한다. CSV는 이 검토를 거쳐 공유한 명시적 mode_id를 사용한다. 가까운 neff 자동 정렬은 불허한다. correlation >=0.8은 screening 정책이며 독립적인 모드 식별 검토를 대체하지 않는다.

convergence JSON: tolerances.Re_absolute/Im_relative/Im_absolute와 axes.mesh/window/pml_thickness/pml_strength/pml_order. 각 axis는 3개 이상의 distinct level을 포함한다. 각 row는 level, result 파일 사양, input_sha256, mode_indices를 가진다. result는 실제 input과 candidates(또는 modes)를 보존한 JSON이어야 한다. mode에는 복소 neff, solver status, 고정 probe의 tracking_fields(E)가 있어야 한다.

gate는 원 결과에서 centroid Re/Im와 부분공간 overlap을 다시 계산하고, 입력 hash 및 해당 axis 외 입력이 동일함을 검사한다. 마지막 두 변화량을 평가한다. 적어도 3수준, 양수 leakage Im, 동일 모드 부분공간이 필요하다. 과거 형상에서 쓰던 5% 수치를 무조건 새 형상 인증에 적용하지 않고 외부 package에서 허용값을 명시한다. FEM 잔차는 이 수렴 검사를 대체하지 않는다.

## V3 정규화 계약

raw_bent_csv 표준 열: wavelength_nm, mode_id, raw_real, raw_imag, r_average_m, r_power_m. COMSOL export에서 이 열로 옮긴 기록과 원 모델을 함께 보존한다. normalization_manifest는 representation, length_unit='m', imag_sign=+1/-1, comparison_radius_basis='geometry'/'power'를 가진다. representation은 bend_reference.py의 명시적 값이다. raw eigenvalue와 이미 정규화된 ewfd2.neff를 구별한다.

gate는 raw 값과 실제 두 반경으로 정규화한 결과가 reference_csv와 일치하는지 검사한다. formulation_review에는 reviewed_by와 equivalent_comparison=true의 근거가 있어야 한다. conformal 근사와 axisymmetric의 정식화 차이가 미해결이면 V3는 NEEDS_EVIDENCE다.

## V6 / V7

- V6: FSM cutoff 2–6, 1200/1800 nm, FM 메시 3수준 및 별도 ring/window probe를 수행한다. FSM 및 FM 마지막 두 절대 neff 변화 <=1e-5는 **계산 전에 선언한 공학적 진단 기준**이며 문헌의 보장 오차가 아니다. finite VPCF는 실제 nFM와 주기 Maxwell nFSM으로 계산한다. HOM cutoff, 정량 문헌 곡선/불확도, 독립 전단면 수렴 증거가 없으면 PASS 불가. d/Lambda=.6 두 파장으로 ESM을 인증하지 않는다.
- V6 후속 qualification: identity, convergence, HOM_reference_comparison 파일 사양. 마지막 JSON은 reviewed_by, cutoff_bracket, primary_reference_data, HOM_field_evidence, reference_V_relative_tolerance를 포함한다. primary_reference_data는 rows 목록에 wavelength_nm, pitch_um, diameter_ratio, V_PCF, uncertainty_absolute를 명시한다. gate는 실제 FM/FSM에서 V를 다시 계산해 기준값·불확도와 비교하며, 사람이 입력한 reference_V_relative_error=0을 신뢰하지 않는다. 아직 없는 자료는 null로 남긴다.
- V7: 원본 7파장과 실패 간격의 최대 1회 중간점 탐색. seed는 capillary 근사이며 기준답이 아니다. 여러 target으로 탐색하되 실제로 실행한 target과 재사용 계산 출처를 남긴다. min core Et proxy=.25, subspace correlation=.8, pair neff 분리<=1e-4, collinearity<=.98 및 spatial overlap>=.85는 후보 screening이며 FM 인증이 아니다.
- branch 결측·단절·끝점 최대값·음수/비정상 손실은 유효 peak가 아니다. 유효 내부 peak를 찾더라도 독립 identity 및 mesh/PML/window convergence, 3개 이상의 spectral refinement peak 수렴이 있어야 한다. ARROW 대비 공진 위치 오차 <=5%는 기존 선언된 공학 기준이며 논문 불확도가 아니다. spectral refinement의 허용오차는 자료와 함께 명시한다. 극소 절대 손실이나 실제 상용 HCF 인증은 별도다.
- V7 spectral_refinement JSON은 relative_tolerance와 levels 목록을 가진다. 각 level의 spectrum 파일 사양은 실제 cases 스펙트럼을 가리킨다. gate가 각 내부 peak를 다시 찾고 peak 주변 파장 간격이 실제로 감소하는지 검사한다. 입력된 peak 위치 목록만으로 통과시키지 않는다.

## 실행 및 종료코드

`python physics_gate.py --evidence-root ../../validation_report/physics_20261121 --baseline ../../validation_report/step7_v089/validation.json --out ../../validation_report/PHYSICS_GATE_20261121.json`

위 명령은 판정 보고서를 생성한다. `--require-all`을 추가하면 미완료 항목이 있는 경우 exit 1이다. 기존 software CI의 exit 정책은 보존한다. 별도 physics workflow가 이 strict gate를 검사하고 실패 결과도 artifact로 보존한다. 이 평가는 새 V2/V3/V6/V7 근거와 과거 scoped PASS의 결합이며 V1–V9 전체를 새로 FEM 재계산한 것으로 표시하지 않는다.
