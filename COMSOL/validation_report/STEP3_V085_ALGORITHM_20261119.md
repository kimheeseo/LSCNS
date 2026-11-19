# STEP 3 v0.8.5 — 알고리즘·단위·출처

## 기존 모듈 위의 추가 구조

| 파일 | 역할 |
| --- | --- |
| `vector_fem.py` | 기존 Maxwell 블록 조립, Nédélec/H1 1·2차 선택, PML ε/μ 텐서, E/H 복원 |
| `vector_spectrum.py` | β² 및 기존 companion pencil의 sparse shift-invert |
| `vector_solver_v085.py` | 잔차/전력/유한성/모드 독립성/분류·입력 계약 |
| `modal_diagnostics.py` | 고정 FE probe, 가중 장 rank, principal angle 부분공간 비교 |
| `modal_study.py` | 재메시·창·PML 매개변수 변화와 근제로 판정 |
| `step3_reference.py` / `validate_step3.py` | FEM과 별개의 원형 Bessel/Hankel 경계 연속조건·검증 실행 |
| `research_step3_v085.mjs` | 실제 Python JSON을 읽는 웹 패널; 벡터 계산 없음 |

## 혼합 정식화와 pencil

전파/시간 관례는 E(x,y,z,t)=E(x,y)exp(iβz−iωt), 따라서 수동 손실은 Im(neff)>0. µm 좌표, k0=2π/(λ_nm·1e-3), β=k0 neff를 사용한다. 횡방향 Nédélec edge 공간은 접선 연속성을 유지하고 종방향 H1 공간과 결합해 일반 절점 3성분 discretization의 스퓨리어스 위험을 줄인다. 이것만으로 모든 가짜 모드 부재가 증명되는 것은 아니므로 Gauss/장 rank/PML/독립 해 검사를 함께 수행한다.

기존 혼합 블록은 다음과 같다.

```
K(β) = [ Ktt + β² Mtt    β Ktz ]
       [ β Kzt           Kzz   ]
```

Et=e, Ez=iβφ를 대입하고 종방향 식을 iβ로 나누면 다음 동치 pencil을 얻는다(β=0 branch는 지원하지 않음).

```
A = [-Ktt   0]        B = [ Mtt    i Ktz]
    [  0    0]            [-i Kzt   Kzz ]
A x = β² B x,   x=(e,φ)
```

B를 직접 역행렬로 만들지 않는다. SciPy SuperLU로 (A−σB)를 분해하고 `(A−σB)^−1 B` LinearOperator에 ARPACK을 적용한다. σ=(k0·target_neff)². 복소/비에르미트/특이 B를 허용한다. 동일 QEP의 기존 2배 크기 companion pencil을 보존해 교차 검사한다. 두 pencil은 독립 물리 기준이 아니므로 별도 정확한 원형 해가 필요하다. **허수부가 정확히 0인 경우에만** 실수 행렬 경로를 사용하며 극소 비제로 흡수를 임의 절삭하지 않는다.

scikit-fem 11–12, SciPy 1.x, Gmsh 4.x를 사용한다. 기본 N1/P1, 정밀 비교 예시는 N2/P2. 기본 혼합 DOF 상한 80,000, 명시적 설정 최대 250,000(메모리 보장 아님), benchmark 120,000. companion 행렬은 혼합 DOF의 2배다. 조립 전 예산 검사와 변환 후 검사로 큰 문제를 막는다. distributed SLEPc/HPC는 TODO이다.

## PML·SBC·H

원형 PML에서 t=(r−R)/L, sr=1+i·strength·t^order, st=1+i·strength·L·t^(order+1)/((order+1)r). 원형 radial/tangential Jacobian으로 ε와 µ를 변환하며 **H=(µ_r)^−1 curl(E)/(i k0 Z0)**에도 변환 역투자율을 적용한다. 외곽은 PEC. 다항식 차수/두께/강도·계산창·메시를 실제 재계산하여 민감도를 검사한다.

SBC는 target neff를 반영한 1차 Robin 근사다. 선택적 `sbc_iterations`(1–6)에서 계산 β로 경계를 갱신하고 이력을 출력한다. COMSOL의 원통형 2차 SBC 또는 정확한 비국소 DtN과 동일하지 않다. 수렴했다고 방사 손실 정확성을 인증하지 않는다.

전력은 PML을 제외한 FE Gauss 적분 `0.5 Re(Ex conj(Hy)−Ey conj(Hx))`에 µm²→m²인 1e-12를 곱한 값이다. 양의 전력을 확인한 후 1 W와 고정 전역 위상으로 정규화하고 실제 정규화 오차를 출력한다. 웹 장은 요소 평균값을 삼각형 중심에 표시한다. 선택적 full quadrature는 실제 FE 적분점이며 요소 평균이나 centroid를 정확한 절점값으로 오인하지 않는다.

## 퇴화·추적·분류

고유값이 가깝다는 이유로 두 편광을 하나로 삭제하지 않는다. 물리영역 고정 polar Gauss probe에서 3성분 E를 FE 보간하고 가중 Gram rank로 동일 위상/배수 장만 제거한다. 재메시·계산창 변경에서도 같은 점/가중치를 유지한다. 퇴화된 쌍은 기저 회전에 불변인 부분공간 principal correlations를 사용하며 최솟값 0.98 이상이 이번 추적 정책이다. 일반 study는 target 근방 첫 두 모드 부분공간을 추적한다(임의 다중 분기 전체 모드 추적은 TODO).

후보 탐색은 0<Re(neff)<3.5 및 |Im(neff)|≤0.5로 제한한다. PML 강도분율>0.4, Gauss 진단>1e-5, noise floor보다 큰 음의 Im은 불안정 후보로 표시한다. 이는 공개된 탐색/분류 정책이며 보편적인 모드 존재 조건이나 에너지 확률이 아니다.

각 모드 QEP 및 pencil 상대잔차, interior H1 test 공간의 약형 Gauss dual norm, 6성분 유한성, 전력, PML 강도분율, x/y mirror 상관, 적분 Stokes, 종방향 장 분율을 기록한다. TE-like/TM-like의 1e-6 분율 기준은 수치 휴리스틱이고 C6v 기약표현 또는 자동 HE/LP 이름 판정이 아니다.

study는 각 개별 모드 Im 부호를 보존한다. 1e-10 기본 imaginary noise floor는 **사용자 조절 수치 정책이며 측정된 불확도가 아니다**. 근제로인 손실을 PASS나 0 dB/km로 인증하지 않는다. 안정성만으로 loss_certified=true를 만들지 않는다.

## 재료·상수 출처와 TODO

| 항목 | 처리 | 출처 / 상태 |
| --- | --- | --- |
| 순수 실리카 | 기존 Malitson B,C Sellmeier 유지 | https://doi.org/10.1364/JOSA.55.001205 |
| Ge/F 도핑 실리카 | 모델 키는 있으나 계산 차단 | TODO: 조성·농도·온도·유효 파장별 검증 계수 원문 필요 |
| 공기 | 기존 n=1.00027 유지, 조건 미지 가정임을 provenance에 표시 | TODO: 온도/압력/습도 및 Ciddor/Edlén 근거; 정밀 해석 입력으로 자동 인증하지 않음 |
| 사용자 복소 n | 명시적 constant n+iκ와 source_url | 실측표/공식 dispersion 계수는 사용자 제공 필요 |
| Z0 | 기존 376.730313668 Ω 보존 | NIST CODATA 2018 https://tsapps.nist.gov/publication/get_pdf.cfm?pub_id=932441 ; 최신 2022는 376.730313412(59) Ω https://physics.nist.gov/cuu/pdf/all.pdf. 상대차 약 6.8e-10. 고유값 조립에는 사용되지 않고 H/전력 정규화에만 사용. |

재료 계수나 임의 물리 보정값을 새로 추정하지 않았다. 합성 기준 시험의 상수 굴절률은 시험 입력이며 상용 섬유 재료 상수가 아니다.

## 공식 1차 문서

- DOLFINx, *Electromagnetic modal analysis for a waveguide*: https://docs.fenicsproject.org/dolfinx/v0.10.0.post2/python/demos/demo_half_loaded_waveguide.html — Nédélec/Lagrange 혼합 및 β² 일반화 pencil의 공식 예제. 본 코드는 scikit-fem 조립이며 이 DOLFINx 코드를 실행하지 않는다.
- scikit-fem API: https://scikit-fem.readthedocs.io/en/stable/api.html — element/Basis/asm/probes API.
- COMSOL 6.4, *Leaky Modes in a Microstructured Optical Fiber*: https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html — 복소 neff·퇴화 HE11·PML/SBC 비교. 공식 표의 부호는 본 전파/시간 관례와 다르다. 공식 모델 수치 재현은 이번 범위 밖이며 합성 원형 해 검증을 해당 PCF의 재현으로 표시하지 않는다.
