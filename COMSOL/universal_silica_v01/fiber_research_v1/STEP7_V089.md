# STEP7 v0.8.9 — V1–V9 검증 스위트

2026-11-20. V1–V9는 **PASS / FAIL / NEEDS_EVIDENCE(근거 필요)** 세 상태다. 전체 항목이 PASS가 아니면 `all_required_passed=false`, `completion_label=NOT_COMPLETE_UNTIL_ALL_V1_V9_PASS`다. CI 성공은 소프트웨어의 실행·판정 동작을 확인하며 물리 검증의 전체 통과를 뜻하지 않는다. STEP8 구조 변경은 시작하지 않는다.

## 검증 정의와 모듈

| ID | 판정 대상 | 근거 / 수용 조건 |
|---|---|---|
| V1 | step-index HE11 b–V 해석해와 실제 혼합 FEM | V=1.8,2.0,2.2,2.35, nc=1.5/ncl=1.44/a=1.3µm 합성 시험. exact J1/K1 Maxwell interface characteristic. neff 상대 오차 ≤10 ppm. b=(neff²−ncl²)/(nc²−ncl²). 기본 모드의 이 구간만 검증하며 모든 고차모드/전 V 구간 인증 아님 |
| V2 | COMSOL microstructured 예제 모드별 복소 neff | 공식 Table1의 10개 PML neff 확인. 원본 parameters.txt의 정확한 입력 미확인 → 근거 필요. 공식 예제 수치와 임의로 비슷한 형상을 만들어 통과시키지 않음 |
| V3 | COMSOL step_index_fiber_bend | 공식 λ=1550nm,ncore=1.4457,nclad=1.4378,a=8µm,R=3mm 등 확인. 2D axisymmetric의 neff×r0 정규화와 현재 conformal 모델 차이, full-precision bent complex neff 미확인 → 근거 필요 |
| V4 | PML 두께·강도·차수 독립성 | 같은 FE 부분공간으로 추적된 실제 strongly radiating circular test. 각 파라미터 3값, max/min Im 차이 / 평균 |Im| ≤5%, min Im>1e−10. 극소 HCF loss 인증 아님 |
| V5 | 실제 요소 크기 vs Re/Im 수렴 | 동일 물리 모드, 4/6/8 elements per material wavelength. 실제 edge_min/max·triangles와 neff 기록. 마지막 centroid ΔRe≤1e−5, radiative centroid relative ΔIm≤5%. 무손실 guided Im≈0은 누설 검증 대상에서 제외 |
| V6 | 고체 코어 PCF / 실제 FSM 기반 판단 | Mortensen et al. (2003)의 n=1.444/1, triangular one-hole defect, V_PCF=(2πΛ/λ)√(nFM²−nFSM²), cutoff π. Γ-cell vector H PWE의 β root로 fixed-λ FSM 계산, cutoff 2/3/4 변화 기록. finite FEM FM 후보와 비교. 수렴·HOM·문헌 정량값 확인 전 ESM PASS 금지 |
| V7 | single-ring ARF의 FEM 공진 vs ARROW | 실제 7λ 스펙트럼·고정 FE probe·2차원 부분공간 및 코어 intensity screen. scalar capillary target은 탐색 seed일 뿐 검증 기준 아님. 내부 FEM loss peak와 ARROW 오차 ≤5%를 정책으로 정하되 메시/PML·동일 FM·peak refinement 증거 없이는 PASS 금지 |
| V8 | 단위·공식 | 독립 SI dB/m↔dB/km, negative Im 보존, 양의 n curvature→음의 D, Gaussian Aeff=πw² |
| V9 | COMSOL CSV 가져오기 / 오차표 | synthetic parser tests. 파일 부호·단위·mode key 명시, unmatched/ambiguous는 근거 필요. CSV 수치 일치만으로 V2/V3나 전체 물리 PASS를 변경하지 않음 |

10ppm/5%/절대오차 등은 **명시한 공학적 수치 수용 정책**이며 논문/COMSOL이 보증한 불확도가 아니다. V2/V3/V6는 알려지지 않은 문헌 오차나 값을 추정해 채우지 않으며 허용값 미정은 null이다.

`validate_step7_v089.py`: 전 항목 조합/판정·출처 hash·환경·시각·gate. `validation_photonics_v089.py`: actual b–V/FSM/ARF 계산. `validation_reference_v089.py`: 확인한 공개 reference facts와 누락 데이터. `validation_csv_v089.py` / 웹 `validation_csv_v089.mjs`: 각각 Python/브라우저의 엄격한 parser 및 오차 계산. 물리식은 기존 postprocess를 호출하며 표시 모듈 `research_validation_v089.mjs`와 분리한다. STEP8의 전체 디렉터리 이관/공식 registry는 이번 단계에 포함하지 않는다.

기존 STEP3 v0.8.5 보고서의 V4는 잔차/장 무결성, V5는 종합 PML/메시 검사였다. 기존 기록은 변경하지 않고, STEP7에서는 **V4=PML 독립성, V5=메시 수렴**으로 사용자 정의에 맞춰 다시 판정한다. 이전 PASS_SCOPED를 항목 이름만 바꿔 복사하지 않는다. fresh STEP3 실제 재실행에서 mode tracking과 sensitivity data를 얻고, 메시를 같은 입력 hash로 재생성해 실제 edge와 요소 수를 대조한다.

## 실행

```bash
python -m pip install -r requirements.txt
OPENBLAS_NUM_THREADS=1 OMP_NUM_THREADS=1 python validate_step3.py --out out_step3
OPENBLAS_NUM_THREADS=1 OMP_NUM_THREADS=1 python validate_step7_v089.py --step3-report out_step3/step3_validation.json --out out_step7
python -m unittest discover -s tests -v
```

V2/V3/V6 근거 필요, V7 후보 모드 검증 실패가 존재해도 결과를 숨기지 않고 저장한다. CLI exit=0은 필수 소프트웨어/단위와 선언된 V1/V4/V5/V8/V9 수치 시험 실행 성공에 대한 값이며 `all_required_passed`를 대신하지 않는다. 외부 물리 항목 전체를 gate로 쓰려면 JSON의 `all_required_passed`를 검사해야 한다. `--skip-photonics`는 계산을 생략해 근거 필요로 두며 cached PASS를 만들지 않는다.

## COMSOL CSV 형식 및 비교

다음 세 형식을 지원한다.

```csv
% lambda (um),mode_id,real(ewfd.neff) (1),imag(ewfd.neff) (1)
1.55,HE11_a,1.445,-3.0E-8
```

```csv
wavelength_nm,mode_id,neff
1550,HE11_a,1.445-3.0E-8i
```

```csv
% ewfd.neff (1)
1.445-3.0E-8i
1.440-5.0E-7i
```

위 숫자는 parser 설명용 합성 fixture이며 공식 V2 재현 데이터가 아니다. comma/semicolon/tab, `%`/`#` 주석, BOM, D 지수, complex a±bi/j, quoted cells를 처리한다. 실수 전용 열에서 Im=0을 추정하지 않는다. complex literal이 순실수인 경우의 Im=0은 파일에 해당 값이 명시된 것으로 처리한다. 단일 열/단일 파장 파일은 사용자가 wavelength와 row-order mode mapping을 확인해야 한다. 레이블로 매칭하며 neff 값이 가까운 행을 임의 재배열하지 않는다. 물리 모드 대응은 COMSOL field export/편광/overlap 근거가 별도로 필요하다.

우리 convention: exp(iβz−iωt), passive Im>0. 반대 시간 관례이면 명시적으로 −1을 곱한다. abs(Im)를 사용하지 않는다. 부호를 선택하지 않으면 비교를 거절한다. 파장 단위 nm/um/m를 헤더 또는 입력으로 확인하며 λ 없는 파일은 기본 파장을 직접 입력한다.

```bash
python validation_csv_v089.py calculated.csv comsol.csv --calculated-imag-sign 1 --reference-imag-sign -1 --re-ppm 10 --im-relative .05 --im-absolute 0 --out error_table.json
python validation_csv_v089.py mode_result.json comsol.csv --calculated-imag-sign 1 --reference-imag-sign -1 --wavelength-nm 1550 --confirm-row-mode-ids --re-ppm 10 --im-relative .05 --im-absolute 0 --out error_table.json
```

웹은 검증 탭에서 STEP3에 가져온 **현재 실제 벡터 결과 사용** 또는 계산 CSV를 선택한다. COMSOL 기준 CSV를 가져오고 부호·단위·모드 대응이 설정돼 있으면 자동 오차표를 만든다. 버튼으로 재계산하거나 JSON 저장도 가능하다. Re ppm=10⁶|차이|/|기준|, Im 기준은 |차이|≤절대 허용+상대 허용×|기준|. 기준 Im=0이면 상대 오차는 null이며 절대 허용으로 판단한다. 입력/단위/부호/오차/설계 변경은 이전 오차표를 무효화한다. 현재 벡터에서 유래한 계산 데이터는 설계 변경 시 함께 제거한다.

## 출처와 한계

- COMSOL 6.4, [Leaky Modes in a Microstructured Optical Fiber](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html), Table1 및 외부 parameters.txt 로드 절차. 참고 논문 Uranus & Hoekstra, [DOI 10.1364/OPEX.12.002795](https://doi.org/10.1364/OPEX.12.002795). 정확한 parameter file 숫자는 미확인.
- COMSOL 6.4, [Step-Index Fiber Bend](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.step_index_fiber_bend/step_index_fiber_bend.html). 축대칭 r0 scaling과 직선 neff 반올림값을 복소 bent reference로 대체하지 않음.
- Mortensen et al., [Modal cut-off and the V parameter in photonic crystal fibers (2003)](https://arxiv.org/pdf/physics/0307010), Eq4/5 및 n=1.444/1 조건. finite sampled V<π는 모든 λ의 ESM 증명이 아님.
- Huang et al., [A Method to Process Hollow-Core Anti-Resonant Fibers into Fiber Filters (2018)](https://doi.org/10.3390/fib6040089), ARROW resonance 관계. 단순 식을 자체 계산해 맞았다고 하는 순환 검증은 하지 않음.

실제 합성 HCF의 core intensity/부분공간 후보 식별 실패는 FAIL로 기록하며 제품 공진 불일치나 ARROW 이론의 반증으로 해석하지 않는다. complete ESM, 극소 confinement loss, PM 재료 상수 또는 정량 접속손실은 이번 PASS 항목의 적용 범위가 아니다. 다음 제안은 V2 원 파라미터 확보, V3 정규화된 COMSOL CSV 확보, V6 converged FM/HOM/FSM 및 V7 모드 식별·PML/메시·공진 스캔 개선이다. STEP8 진행 여부는 해당 결과를 검토한 뒤 판단한다.
