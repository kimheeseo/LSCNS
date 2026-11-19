# LS Wave Optics — STEP 0 COMSOL 광섬유 기능 재감사
기준일: 2026-11-19 (KST)  
대상: kimheeseo/LSCNS / COMSOL/universal_silica_v01  
성격: 승인 검토용 기술 감사. 이번 작업에서 앱 코드·버전·커밋·배포 변경 없음.

## 1. 결론과 감사 범위

**현재 목표의 출발점은 “v0.7.1에 없는 기능을 새로 작성”이 아니라 “이미 추가된 v0.8.x 연구 모듈의 물리·검증 상태를 정리하고 필요한 부분을 보완”이다.** 기존 동작을 유지하는 추가 방식은 그대로 적용한다. 요청한 승인 경계에 따라 이번에는 STEP 0만 수행했다.

- 고정 검토 스냅샷: [main 3967bd334091](https://github.com/kimheeseo/LSCNS/tree/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL). 재현 가능한 근거를 위해 아래 코드 링크는 모두 이 커밋으로 고정했다. 저장소 HEAD의 시간은 하위 프로젝트 수정 시간과 구분한다.
- 웹 scalar 헤더/solver는 **v0.7.1**이며, 별도 연구 모듈에는 **v0.8.0–v0.8.3**이 함께 있다. 하나의 버전으로 이미 통합됐다고 판정하지 않는다.
- 저장소에는 2026-11-18 STEP 0 감사와 후속 단계 문서도 존재한다. 이는 기존 자산이며, 이번 요청의 “새 STEP 0 승인 전 진행 금지”를 대체하지 않는다.
- 이번 확인은 소스·설정·저장된 검증 문서의 정적 감사와 공식 문서 열람이다. COMSOL .mph 실행, Python FEM 실행, 브라우저 동작 재측정은 하지 않았다. 문서에 있는 과거 PASS/FAIL은 신규 시험 성적으로 취급하지 않았다.
- COMSOL 공식 문서 6.4 User’s Guide/HTML 및 Application Library를 우선 사용했다. Fiber Simulator는 실제 열람한 **6.3** PDF임을 표시했다. “최신 COMSOL 버전 전체 기능” 감사라는 주장은 하지 않는다.
- “확인”은 문서나 코드에서 확인한 내용, “미확인”은 이번 자료로 확인하지 못한 사실, “미검증”은 구현이 있어도 물리 정확도가 입증되지 않은 상태다.
- 난이도·필요 알고리즘·검증 방법·우선순위는 **우리 툴에 대한 엔지니어링 제안**이다. COMSOL 공식 사양과 구분한다.

## 2. 저장소 구조·파일 맵

COMSOL/ 상위에는 COMSOL_Install/, research/, public_reference_benchmarks/, g657a1_validation/, validation_report/, universal_silica_v01/ 및 실행 스크립트가 있다. 본 감사의 구현 상세 범위는 지정한 universal_silica_v01과 관련 validation_report 문서다. 다른 연구 폴더 전체에 대해 “기능이 없다”고 판정하지 않았다.

| 계층 | 확인한 주요 파일 | 역할·상태 |
| --- | --- | --- |
| 기존 웹 진입 | [web/index.html](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/web/index.html); [web/solver.mjs](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/web/solver.mjs) | 헤더·기준 scalar VERSION은 v0.7.1; 원형 약유도 radial P1 FEM 보존 |
| 기존 웹 벡터 | [web/vector_fdfd_browser.mjs](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/web/vector_fdfd_browser.mjs); [web/yee_solver.mjs](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/web/yee_solver.mjs); [web/vector_geometry.mjs](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/web/vector_geometry.mjs) | collocated H-field FDFD와 별도 staggered Yee 경로가 이미 존재. Python FEM과 구분 |
| 설계·결과 웹 연결 | [web/research_fiber_v1.mjs](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/web/research_fiber_v1.mjs); [web/research_step5_v083.mjs](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/web/research_step5_v083.mjs) | A–H JSON 편집/Python 결과 import; Bloch도 Python JSON 시각화. index.html이 v0.8.x 연구 모듈 로드 |
| Python 기준·기존 연구 | [universal_radial.py](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/universal_radial.py); [full_vector/vector_modes.py](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/full_vector/vector_modes.py); [hcf_complex_fem/hcf_complex_fem.py](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/hcf_complex_fem/hcf_complex_fem.py) | radial scalar / PML 없는 벡터 FDFD / HCF 복소 scalar FEM을 각각 보존 |
| 프리셋·입력 계약 | [fiber_research_v1/presets.json](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/fiber_research_v1/presets.json); [fiber_research_v1/fiber_design.schema.json](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/fiber_research_v1/fiber_design.schema.json); [fiber_research_v1/schema.py](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/fiber_research_v1/schema.py) | schema_version=1.0.0; 실제 JSON 25개: A5/B5/C2/D1/E5/F1/G3/H3 |
| 기하·벡터 코어 | [fiber_research_v1/geometry.py](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/fiber_research_v1/geometry.py); [fiber_research_v1/vector_fem.py](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/fiber_research_v1/vector_fem.py); [fiber_research_v1/materials.py](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/fiber_research_v1/materials.py) | gmsh OCC·원형 PML·scikit-fem Nédélec/H1·SciPy sparse QEP. 새로 만들 대상이 아니라 검증/보완 대상 |
| 특성·Bloch | [fiber_research_v1/physical_metrics.py](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/fiber_research_v1/physical_metrics.py); [fiber_research_v1/step5_bloch.py](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/fiber_research_v1/step5_bloch.py) | v0.8.0 후처리; v0.8.3 스칼라 fixed-β Bloch. full-vector periodic Maxwell은 미구현 |
| 굽힘·응력 | [fiber_research_v1/bending.py](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/fiber_research_v1/bending.py); [fiber_research_v1/stress_optic.py](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/fiber_research_v1/stress_optic.py) | v0.8.1 등가 굽힘·v0.8.2 generalized-plane-strain/광탄성 후보 경로. 제품 정량값 미인증 |
| 문서·회귀 자산 | [web/CHANGELOG.md](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/web/CHANGELOG.md); [web/M3_M6_YEE_VALIDATION_20261112.md](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/web/M3_M6_YEE_VALIDATION_20261112.md); [fiber_research_v1/README.md](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/fiber_research_v1/README.md) | 기능 추가 이력·실험 결과·제약이 공존. 오래된 설명과 새 구현 상태의 불일치 존재 |

의존 라이브러리 선언은 [fiber_research_v1/requirements.txt](https://github.com/kimheeseo/LSCNS/blob/3967bd3340914becc3079248a9326d70060c9bf5/COMSOL/universal_silica_v01/fiber_research_v1/requirements.txt)에 numpy>=1.26,<3; scipy>=1.12,<2; scikit-fem>=11,<13; gmsh>=4.12,<5; jsonschema>=4.20,<5로 되어 있다. 이것은 허용 범위이며 이번 환경의 설치·실행 버전 증거는 아니다. 후속 재현 시험에서는 실제 설치 버전과 환경을 고정해야 한다.

## 3. 요청의 “없는 것”과 현재 소스의 차이

| 대상 | 현재 확인한 구현 | 감사 판정 |
| --- | --- | --- |
| 웹/로컬 일반 벡터 모드 | 웹 Yee-FDFD와 Python HCurl/H1 FEM이 이미 존재 | 기능 부재로 볼 수 없음. 독립 HE11·수렴 검증 완료와는 별개 |
| 임의 2D/PCF 단면 | 기존 웹 격자 단면과 새 gmsh 프리셋 단면 | 벡터 솔버 연계 코드 있음. 모든 CAD 변형의 정확한 형상 인증은 미완료 |
| 누설손실 직접 계산 | 복소 neff → 부호 보존 dB/m·dB/km | 정량 예측 인증 미완료. 재료흡수가 있으면 총손실이며 순수 confinement loss와 구분 |
| 굽힘 | 기존 근사 외에 Python v0.8.1 n_eq 경로 | COMSOL axisymmetric bend와 정량 동치 미검증 |
| 복굴절/응력 | generalized-plane-strain + 광탄성 + anisotropic 후보 | 일반 3×3 ε 텐서 전체 지원 아님; 실제 Panda 물성·모달 B 검증 필요 |
| MCF 실계산 | N코어 단면을 벡터 코어에 넣는 경로 | supermode 정밀도 및 실제 케이블 random XT 검증 별도 |
| Bloch 밴드 | v0.8.3 scalar fixed-β FD 고유계 | scalar proxy만 있음. 일반 PCF full-vector projected band 검증 미완료 |
| overlap 접속손실 | 스윕용 field overlap는 있음 | 검토한 연결 경로에서 일반 E/H 접속손실 기능은 확인 못함. 스윕 추적 overlap와 접속 효율을 혼동하면 안 됨 |

새 Python FEM의 현재 출력은 validation.level=EXPERIMENTAL_NOT_BENCHMARKED, mesh_converged=false, pml_converged=false, loss_dB_km_certified=false로 고정돼 있다. 이는 “이미 수렴이 입증됐다”는 플래그가 아니다.

## 4. COMSOL 광섬유 연구 기능 매트릭스

EWFD = Electromagnetic Waves, Frequency Domain. EWBE = Electromagnetic Waves, Beam Envelopes.  
P0 = 수치 신뢰성 기반, P1 = 핵심 연구 기능, P2 = 검증 기반 확보 뒤 다중물리·최적화.  
난이도는 초기 코드 작성만 아니라 범용 연구용 검증까지 포함한다.

| 기능 | COMSOL 구현 방식: 인터페이스·경계조건·연구 유형 | 광섬유 연구 용도 | 난이도 | 필요 알고리즘: 제안 | 검증 방법: 제안 | 우선순위 | 근거 문서 제목 / URL |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 2D 벡터 모드·복소 neff | EWFD + 2D Mode Analysis; 주어진 주파수에서 전파상수 고유계. 개방 단면은 PML/SBC | PCF·HCF·MCF의 유도/누설 모드와 6성분 장 | 상 | 기존 HCurl(E_t)+H1(E_z) β-이차 고유계 재사용; sparse shift-invert; H 복원 | 원형 HE11 정확 벡터 기준해, 약유도 LP01 별도 비교, 메시/계산창 수렴 | P0 | [Mode Analysis and Boundary Mode Analysis (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.woptics/woptics_ug_modeling.5.46.html); [Leaky Modes in a Microstructured Optical Fiber (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html) |
| PML·직접 누설손실 | EWFD + 좌표 변환 PML 도메인 + Mode Analysis; MOF 예제는 cylindrical PML, 횡방향 파장 지정 | 유한 클래딩 방사손실과 복소 neff | 상 | 현재 radial complex stretching의 ε/μ 텐서 검증; 물리창·PML·메시 분리 실험 | 두께/강도/차수/창/벽 해상도 수렴; 부호·손실 바닥 검증; 독립 개방계 비교 | P0 | [Leaky Modes in a Microstructured Optical Fiber (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html) |
| Scattering BC | EWFD SBC의 파형/1·2차 선택, Mode Analysis의 전파상수 차감 옵션. MOF는 cylindrical wave | PML 대안 및 외부 방사 경계 교차 확인 | 상 | β 의존 방사 임피던스; 현재 fixed-target 1차 Robin 근사 범위 평가/보완 | 동일 단면 PML/SBC·창 크기 비교. 실수부 일치만으로 작은 허수부 정확도 판정 금지 | P1 | [Scattering Boundary Condition (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.woptics/woptics_ug_optics.6.29.html); [Leaky Modes in a Microstructured Optical Fiber (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html) |
| 유한 PCF 클래딩·결함 코어 | 기하·재료 단면의 EWFD 모드 해석; 공식 PCF gallery는 index-guided 기본/고차 모드 분산 | Λ, d/Λ, 링 수, 결함·타원 공기공 설계 | 상 | 기존 gmsh 생성기 재사용; 정확한 hole/defect topology, 경계 적합 메시 | 공식 PCF 조건 고정 비교, 링 수·공기공 경계·메시 수렴 | P1 | [Photonic crystal fiber blog model (Application ID 75371)](https://www.comsol.com/model/photonic-crystal-fiber-blog-model-75371) |
| 고유주파수 밴드·Floquet | EWFD + unit-cell Floquet periodic BC + Eigenfrequency + k sweep; 공식 예제는 Ez 편광 축소 모델, 주파수 의존 재료는 nonlinear ARPACK | 주기 클래딩 band/gap; 광섬유 projected band 설계의 기반 | 상 | 대응 경계 자유도 위상 연결; 주어진 βz의 벡터 ω(k) 문제는 별도 정식화. 기존 scalar proxy는 보존 | 균일 매질 dispersion, 주기 경계 일치, k·메시 수렴, 공식 Ez 예제 재현. 완전 갭은 BZ 내부도 확인 | P1 | [Band-Gap Analysis of a Photonic Crystal (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.bandgap_photonic_crystal/bandgap_photonic_crystal.html) |
| 굽힘: 등각 근사/축대칭 | 공식 bend는 직선 2D 모드와 굽힘 2D axisymmetric 방위각 모드+PML. 등각 n(1+x/R) 전용 내장 노드는 미확인 | 반경·방향별 모드 이동, macro-bend loss | 상 | 현재 등가 굴절률 근사와 정확한 곡선좌표/축대칭 해석을 구분; R_eff는 근거 입력 | R→∞ 직선 극한, 공식 축대칭 예제, 같은 모드 추적·PML 수렴 | P1 | [Step-Index Fiber Bend (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.step_index_fiber_bend/step_index_fiber_bend.html) |
| 분산·파장 스윕 | EWFD Mode Analysis + 파장 Parametric Sweep; Fiber Simulator에서 group delay·dispersion 출력 | ng, D, S, ZDW, 편광별 분산 | 중 | 재료 n(λ) + 같은 모드의 neff(λ), 미분·overlap 추적; 재메시 보간 | 파장 간격·피팅 차수·메시를 독립 변화; 모드 전환 검출; 재료/도파로 분산 분리 | P1 | [Fiber Simulator (6.3, PDF pp.4–5)](https://www.comsol.com/model/download/1423651/applications.fiber_simulator.pdf); [Optically Anisotropic Waveguide (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.optically_anisotropic_waveguide/optically_anisotropic_waveguide.html) |
| 형상/재료 복굴절 | EWFD에 비등방 ε tensor + Mode Analysis; 공식 anisotropic 예제는 diagonal/off-diagonal 재료와 mode following | 형상·응력 유래 Δneff, beat length, 편광축 | 상 | 현재 xy 대칭 블록+zz 지원 검토; 일반 xz/yz 텐서는 별도 확장; 퇴화 부분공간 추적 | 원형 퇴화 보존, 전체 구조 회전 불변성, 기준 anisotropic 모드 비교 | P1 | [Optically Anisotropic Waveguide (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.optically_anisotropic_waveguide/optically_anisotropic_waveguide.html) |
| 열응력·응력광학 결합 | Solid Mechanics Stationary → 광탄성 굴절률/ε 변화 → EWFD Mode Analysis. 공식 예제는 Structural Mechanics 추가 필요; generalized plane strain 예제 존재 | Panda·bow-tie 잔류 응력, 열/인장 복굴절 | 상 | 기존 탄성 FEM·열변형률·p11/p12 매핑 재사용; 적용 가능한 strain/stress-optic 규약 관리 | 무응력·균일 열변형 기준, 응력 평형, 기계/광학 메시 별도 수렴, 모달 B 독립 비교 | P2 | [Stress–Optical Effects in a Photonic Waveguide (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.stress_optical/stress_optical.html); [Stress–Optical Effects with Generalized Plane Strain (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.stress_optical_generalized/stress_optical_generalized.html) |
| MCF supermode | 전체 단면 모드해석을 적용. 공식 Directional Coupler는 numeric port/Boundary Mode Analysis 및 EWBE 전파로 even/odd beating | 2/N코어 고유모드, Δβ, 결합길이·전력 전달 | 상 | 기존 N코어 형상+벡터 코어 재사용, 모드 위상/편광/퇴화 정규화, 결합행렬 | even/odd splitting·큰 pitch 극한·전력 보존·대칭성. N코어 광섬유 전용 공식 예제는 미확인 | P1 | [Directional Coupler (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.directional_coupler/directional_coupler.html) |
| Overlap 접속·결합 손실 | Numeric Port/Boundary Mode Analysis 및 E/H overlap. 공식 fiber-to-fiber 예제는 EWBE+렌즈+PML 자유공간 결합 | 모드 불일치·횡오정렬·회전 coupling. 실제 접속 반사/갭은 추가 문제 | 중 | 전력 정규화 E/H overlap, 좌표 이관; 반사까지는 mode matching/전파 해석 | 동일 모드 η=1, Gaussian 이동/크기 해석해, 위상 불변성, 필요한 경우 port S-parameter | P1 | [Directional Coupler (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.directional_coupler/directional_coupler.html); [Single Mode Fiber-to-Fiber Coupling (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.single_mode_fiber_coupling/single_mode_fiber_coupling.html) |
| 파라메트릭 스윕·모드 추적 | Parametric Sweep + Mode Analysis/Eigenfrequency; mode following 공식 예제 | 기하·파장·반경·재료의 설계 지도 | 중 | 기존 CLI/웹 JSON 유지; overlap/퇴화군 추적, 실패점·수렴 지표 포함 | 단일 계산과 sweep 동일점 일치, 단순 정렬과 물리적 추적 구분, 재시작·결과 왕복 | P1 | [Optically Anisotropic Waveguide (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.optically_anisotropic_waveguide/optically_anisotropic_waveguide.html); [Band-Gap Analysis of a Photonic Crystal (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.bandgap_photonic_crystal/bandgap_photonic_crystal.html) |
| 파라미터·형상 최적화 | General Optimization의 목적함수·제어변수·제약; Optimization Module. 공식 PC 예제는 형상 최적화 | D/Aeff/손실/복굴절의 제약 설계 | 상 | 검증된 목적함수 위에 bounded 탐색·캐시; 고급 감도/adjoint는 후속 | 작은 기준 문제·grid 비교·독립 최종 재계산; 수렴 실패점 최적값 배제 | P2 | [The General Optimization Study (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.opt/opt_ug_solver.8.06.html); [Optimization of a Photonic Crystal for Signal Filtering (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.photonic_crystal_filter_optimization/photonic_crystal_filter_optimization.html) |
| 후처리 MFD | Fiber Simulator의 MFD(λ) 출력 확인; 공개 PDF의 Petermann/second-moment 등 정확한 정의는 미확인 | 모드 크기·정렬 공차·접속 비교 | 중 | second-moment x/y MFD, 축대칭 Petermann-II 분리; 선택한 정의와 적분 영역 명시 | Gaussian 기준·크기 scaling·축 회전·적분 해상도; 비대칭 모드 단일 지름 한계 | P1 | [Fiber Simulator (6.3, PDF pp.4–5)](https://www.comsol.com/model/download/1423651/applications.fiber_simulator.pdf) |
| 후처리 Aeff·γ | 장/전력의 Derived Values 및 사용자 수식으로 구성할 설계. Aeff·γ 전용 자동 내장 변수·버튼은 조사 자료에서 미확인 | 비선형성·유리/가스 기여 비교 | 중 | 기존 Et 기반 Aeff 보존; 균질 약유도 γ 근사와 일반 벡터 재료가중 적분 구분; n2 출처 필수 | Gaussian Aeff, field scaling 불변성, PML 제외, 재료·정규화·적분 수렴 | P1 | [Wave Optics Module User’s Guide (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.woptics/WaveOpticsModuleUsersGuide.pdf); [Fiber Simulator (6.3, PDF pp.4–5)](https://www.comsol.com/model/download/1423651/applications.fiber_simulator.pdf) |
| 후처리 모드 분류·가짜 모드 | MOF 예제의 HE/TE/TM-like 장 비교, anisotropic 예제의 편광 및 mode following. 범용 spurious-mode 자동 판정 보증은 미확인 | HOM·편광·대칭 분류와 물리 모드 선별 | 상 | 장 대칭 연산, E/H 구조·Gauss 잔차·PML participation·overlap; eigenvalue만으로 중복 제거 금지 검토 | HE11/TE/TM 기준, 퇴화 모드 유지, 메시/PML 변화에 안정한 부분공간 | P0 | [Leaky Modes in a Microstructured Optical Fiber (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html); [Optically Anisotropic Waveguide (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.optically_anisotropic_waveguide/optically_anisotropic_waveguide.html) |
| 기하·메시·주기 경계 적합 | 공식 band 예제는 대응 주기 경계 메시 복사; MOF/bend는 개방 경계용 별도 PML 도메인 | 얇은 튜브 벽·곡률·재료 경계와 주기 셀 해상도 | 상 | 기존 gmsh physical group·원형 PML 재사용; 벽 법선 실해상도·곡률·국소 λ/n 규칙 추가 설계 | 벽 횡단 요소 실측, 면적/라벨 보존, min/max·종횡비·주기쌍 검사. 모든 섬유의 보편 최소 요소 수는 미확인 | P0 | [Band-Gap Analysis of a Photonic Crystal (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.bandgap_photonic_crystal/bandgap_photonic_crystal.html); [Leaky Modes in a Microstructured Optical Fiber (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html); [Step-Index Fiber Bend (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.step_index_fiber_bend/step_index_fiber_bend.html) |

Aeff·γ 행의 사용자 수식 방식은 당사 구현 설계 제안이며, 공식 전용 버튼이 확인됐다는 뜻이 아니다. 라이선스 추가 의존성이 확인된 것은 위 응력광학 예제의 Structural Mechanics, 최적화의 Optimization Module이다.

## 5. 코드에서 확인한 우선 보완점

아래는 구현 착수 내역이 아니라 승인 후 작업 후보 목록이다. 정적 검토로 찾은 위험과 실제 계산으로 확인한 오류를 구분했다.

| 항목 | 코드 근거 | 관찰·영향 | 후속 단계 제안 |
| --- | --- | --- | --- |
| 벽 두께 메시 기준이 목표값에 머묾 | geometry.py: wall_size=t/through_wall_elements, target_elements_per_wall 보고 | 각 벽 법선 방향의 실제 요소 수를 측정한 인증은 아님. 얇은 벽 손실에 직접 영향 | STEP 2에서 벽 횡단 해상도 검사·미달 시 실패/경고 |
| 실제 재료 n 대신 고정 nmax | geometry.py: nmax=max(1.00027,1.45) | 사용자 정의 고굴절률 등에서 실제 국소 λ/n 제한을 반영하지 못할 수 있음 | 도메인별 재료/파장으로 크기 규칙 산출; bulk 전체 적용 여부 검사 |
| 곡률 규칙의 의미가 다름 | MeshSizeFromCurvature=0; curvature_segments는 Distance sampling에 사용 | 곡률 적응 메시를 보증하는 값이 아님 | 경계 chord 오차·곡률·최소 요소 규칙을 명시하고 실측 |
| C6v 영역 축소 | 전체 도메인 gmsh 생성; 기존 상태 문서에서 미구현 | 계산량 절감 옵션으로 인증되지 않음 | 전체 도메인 유지; 대칭조건은 별도 검증 뒤 선택 기능 |
| 혼합 요소 사용만으로 spurious 제거 보증 불가 | vector_fem.py: ElementTriN1+ElementTriP1, weak Gauss 진단 | Gauss는 현재 진단/후처리 선별. 독립 물리성 검증을 대신하지 않음 | HE11·고차 기준모드·비물리 모드 분리 시험 |
| 퇴화 모드 삭제 위험 | solve(): neff 실·허수 차이 각각 1e-10 미만이면 중복 제외 | 서로 다른 편광의 퇴화 HE11 쌍도 고유값만으로 합쳐질 가능성. 실제 영향은 재현 시험 필요 | 고유벡터 독립성/부분공간 overlap와 함께 중복 판정 |
| 대칭 분류 미완료 | symmetry_class='unclassified ...' | 편광 비율 제공과 HE/TE/TM·군대칭 분류 완성은 다름 | 필드 기반 분류·퇴화군 처리 추가 설계 |
| PML 수렴 실험의 추적 한계 | cli.py: 첫 후보 선택, 결과 note='Nearest-index pairing only' | λ/메시/PML 변화 중 다른 모드를 비교해 안정하다고 판단할 위험. STEP4/6a 별도 overlap 코드를 일반 study와 구분 | 물리영역 field overlap, 재메시 보간 및 모드군 매칭 |
| 메시 sweep의 벽 해상도 고정 가능성 | cli.py study는 elements_per_wavelength 변화; geometry.py는 min(h_active,wall_size) | wall_size가 지배하면 wavelength 설정을 바꿔도 얇은 벽 해상도가 실제로 변하지 않을 수 있음 | 최소 3단계 실제 메시 차이·벽/곡률/벌크 개별 refinement 확인 |
| SBC 근사 | fixed target neff로 kt 계산, 1차 Robin | 고유모드 β를 쓰는 COMSOL 모드용 SBC와 자동 동일시 불가 | 근사 적용범위 평가, 필요시 β 반복/비선형 경계 |
| ε 텐서 범위 | pmxx,pmxy,pmyy,pmzz | PM 단면의 특정 대칭에는 유용하나 εxz/εyz까지 포함한 임의 텐서가 아님 | general tensor 지원 여부를 schema/화면/출력에 명시 |
| provenance 오류 | materials.py의 Malitson URL은 10.1364/AO.4.000901 | 사용 중인 SiO2 Sellmeier 계수의 Malitson 1965 원문 DOI와 불일치 | 출처 메타데이터 교정 후보; 계수를 근거 없이 바꾸지 않음 |
| 재료 가정 | materials.py air n=1.00027, Ge/F 모델 TODO 예외 | 실제 가스 압력/온도/조성 분산 및 도핑 농도 모델과 구분 필요 | 조성별 공식 계수/유효범위/단위가 확보될 때만 연결 |
| 버전·상태 문구 불일치 | web 헤더 v0.7.1, 연구 README v0.8.3; FEM warnings에 옛 stress/Bloch 미계산 문구 | 어떤 엔진/특성값이 검증됐는지 오해 가능 | 후속 변경 시 앱/엔진/schema/검증상태를 분리 표기 |

Malitson의 정확한 원문은 [Malitson, Interspecimen Comparison of the Refractive Index of Fused Silica (1965)](https://doi.org/10.1364/JOSA.55.001205)이다. 기존 웹 M2의 교정된 출처와 새 Python materials.py가 일치하는지 함께 관리할 필요가 있다. 이번에는 출처와 계수 모두 수정하지 않았다.

## 6. 물리 모델·후처리에서 지킬 구분

### 6.1 유한 단면과 주기 클래딩

유한 섬유 모드는 λ 또는 ω를 고정하고 β/neff를 구한다. 주기 셀 밴드는 Bloch k와 필요시 βz를 주고 ω를 구한다. 이 두 고유문제의 경계조건·미지수·결과 스키마는 구분해야 한다.

현재 step5_bloch.py는 -∇t²Ez+β²Ez=(ω/c)²n²Ez의 scalar proxy이며, n은 실수·비분산이다. 일반적인 유한 βz PCF의 hybrid vector Maxwell 문제 전체를 대체하지 못한다. 공식 BAND 예제도 Ez 편광으로 축소한 2D 문제이므로, 해당 특별한 조건을 맞춘 검증과 일반 projected bandgap의 검증은 별개다. 특정 k 경로 또는 거친 BZ 표본에 gap이 있더라도 완전 밴드갭으로 확정하지 않는다.

### 6.2 복소 neff 부호와 손실

현재 Python 규약은 exp(iβz−iωt), ñ=n+iκ다. 따라서 수동 전진 모드의 Im(neff)>0가 감쇠이며

α[dB/m] = (20/ln 10)·(2π/λ[m])·Im(neff),  
α[dB/km] = 1000·α[dB/m].

COMSOL User’s Guide는 exp(iωt−αz), α=δz+iβ, neff=−iα/k0 규약을 설명한다(UG pp.90–91). 같은 수동 전진 모드라도 허수부 부호를 변환해 비교해야 한다. abs(Im neff)로 음수 손실을 숨기면 안 된다. κ가 0이 아니면 출력 손실에는 재료흡수와 방사손실이 함께 포함된다.

COMSOL MOF 예제 자체도 PML과 SBC 사이에서 실수부보다 작은 허수부의 일치가 훨씬 약함을 설명한다. 따라서 한 고유값의 많은 소수 자릿수나 작은 선형대수 잔차만으로 초저 누설손실을 인증하지 않는다.

### 6.3 굽힘·응력

공식 axisymmetric bend는 굽힘 중심 주위의 방위각 전파를 다룬다. 현재 2D 단면 등가 굴절률은 약곡률 근사이며 이 문제와의 정량 동등성은 별도 검증 대상이다. R_eff 보정계수는 특정 섬유·하중·문헌 조건에 종속되므로 보편 상수를 임의 입력하지 않는다.

응력광학의 p11/p12는 무차원 strain-optic 계수이고 stress-optic 계수와 단위·정의가 다르다. ΔT, 탄성, 열팽창, 도핑/조성의 출처가 없으면 실제 Panda 성능을 계산했다고 표시하지 않는다. 국소 굴절률 차이·코어 평균 B와 모달 neff splitting도 구분한다.

### 6.4 MFD·Aeff·γ·접속

같은 “MFD”라도 second-moment, Petermann-II, Gaussian 1/e² 정의가 다르다. 비원형 모드는 축별 값을 보고하고 정의를 결과에 저장한다. Et 기반 Aeff는 해당 정의로 표시한다. 일반 벡터·다중 재료의 비선형 계수는 균질 약유도 γ≈2πn2/(λAeff) 근사와 적용범위가 다르다. n2 및 파장·조성·압력 조건은 제공된 출처로만 입력한다.

field overlap에 의한 무반사 결합 효율과 실제 접속 불연속의 Fresnel 반사·갭·산란은 다른 문제다. 단순 overlap만으로 총 접속손실을 완전히 예측했다고 표현하지 않는다.

## 7. 검증 현황표 — 이번 실행 결과 아님

| 확인 항목 | 근거 | 이번 판정 |
| --- | --- | --- |
| A–H 프리셋 수·스키마 필드 | 실제 presets.json 파싱: 25개, 8계열, 버전 1.0.0 | 구조 존재 확인. 각 제품/형상의 물리 인증 아님 |
| Python mixed FEM·원형 PML·SBC·6성분 출력 | vector_fem.py 정적 검토 | 구현 존재 확인. 새 수치 검증 PASS 부여 안 함 |
| 기존 브라우저 Yee neff mesh convergence | 2026-11-12 저장 보고서: 마지막 N72/N96 차이 약 5.549×10^-6 > 당시 10^-6 기준 | 과거 기준 미통과. 새 Python FEM의 성적과 혼용 금지 |
| 기존 브라우저 직선 PML 손실 | 같은 보고서: 약 −7.743~14.963 dB/km | 과거 부호 불안정·정량 미검증. 이번 재측정 아님 |
| 새 Python FEM HE11/PML 인증 | README·출력 validation 플래그 | 미검증. 과거 small-mesh smoke와 구분 |
| Bloch full-vector·complete bandgap | step5_bloch.py·README | scalar proxy로 한정. full-vector는 미구현 |
| PM 제품 modal B·굽힘 절대손실 | 소스·상태 문서 | 구현 후보 있음, 독립 정량 검증 미완료 |
| 기존 앱 회귀·사용자 브라우저 실행 | 이번에는 실행하지 않음 | 미확인. 코드 변경도 없음 |

요청의 V1/V2/V4/V5/V6/V7 식별자는 이번 메시지에서 정식 시험 정의·오차 허용치가 제공되지 않았다. 확인한 승인 문서에도 본 요청과 일치하는 번호별 완전한 명세를 확인하지 못했다. 임의로 번호 의미를 확정하거나 “V1/V4/V5 통과”를 선언하지 않는다. 후속 검증 계획에서 원문 명세를 연결하거나 번호 매핑 초안을 별도로 승인받아야 한다.

## 8. 미확인·TODO 등록부

| 항목 | 상태 | 필요한 자료 또는 확인 |
| --- | --- | --- |
| COMSOL 전용 등각 굽힘 내장 노드 | 미확인 | 해당 기능명·정식 수식 문서 |
| 임의 βz의 full-vector PCF projected band 전용 공식 절차 | 이번 공식 자료에서 미확인 | 축/파수/고유변수·편광이 명확한 기준 모델 |
| N코어 광섬유 전용 Application Library 표준 모델 | 미확인 | 현재 coupler 일반 원리만 확인. MCF 독립 수치 기준 필요 |
| COMSOL MFD 세부 정의 및 Aeff·γ 자동 내장 변수 | 미확인 | 모델 내부 수식/변수 문서 |
| 보편 spurious-mode 필터와 얇은 벽 최소 요소 수 | 미확인 | 특정 요소 차수·형상·오차 목표에 맞는 수렴 검증 필요 |
| 실제 Ge/F 농도별 Sellmeier | Python 연구 코어 TODO | 조성 단위·계수·유효 파장/온도 범위와 원문 |
| Panda/bow-tie rod 물성 | 실제 제품 정량값 미검증 | E,ν,α,p11,p12,n(λ),ΔT 및 조성·열이력 출처 |
| 초저손실 Im(neff) 분해능 | 미검증 | 메시/창/PML/선형계·고유계 오차의 독립 수렴 및 기준해 |
| V번호별 수용기준 | 미확인 | 번호 정의·입력 파일·비교량·허용오차·기준 데이터 |
| STEP 6c–6e의 상세 범위 | 현재 요청에 세부 정의 없음 | 해당 단계 착수 전에 세부 사양 필요. STEP 0 완료에는 영향 없음 |

## 9. 승인 후 제안 — 다음 한 단계는 STEP 2 보완

요청한 순서 STEP 0 → 2 → 3 → 고체 코어 PCF → 5 → HCF/ARF → 6a → 6b → 6c → 6d → 6e를 유지하되, 이미 있는 코드에는 **보완·검증**으로 적용한다. 승인 후에도 여러 단계를 한 번에 진행하지 않는다.

다음 단일 단계인 STEP 2에서 제안하는 작업:
1. 기존 gmsh/JSON/웹 import 경로를 보존하고 설계 문서 작성.
2. 실제 재료 기반 λ/n·벽 법선 요소 수·곡률 규칙, physical labels·원형 PML·형상 적합성 검사 보완.
3. 실제 크기·종횡비·벽 횡단 요소 수와 3단계 메시 차이를 검증표로 기록.
4. 기하 근사와 C6v 미지원 범위를 한계 보고서에 명시.
5. 기존 웹/Colab/JSON 회귀 확인 후 다음 단계 제안. STEP 3 수치 코어 변경은 별도 승인 전 시작하지 않음.

버전은 v0.8.0으로 되돌리지 않는다. 현재 연구 모듈의 최고 표시가 v0.8.3이므로 **실제 보완 구현 시 v0.8.4를 후보**로 하고, 앱·엔진·schema 버전 및 검증 상태를 분리해 변경 이력을 남기는 것을 제안한다. 이번 감사만으로 버전은 변경하지 않았다.

승인 대상은 이 STEP 0 감사 결과와 “다음은 기존 STEP 2 기하·메시 엔진만 보완한다”는 범위다. 현재 단계에서 후속 코드 구현·커밋·배포는 수행하지 않았다.

## 10. 공식 문서 목록

- UG: [Wave Optics Module User’s Guide (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.woptics/WaveOpticsModuleUsersGuide.pdf)
- MODE: [Mode Analysis and Boundary Mode Analysis (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.woptics/woptics_ug_modeling.5.46.html)
- MOF: [Leaky Modes in a Microstructured Optical Fiber (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html)
- SBC: [Scattering Boundary Condition (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.woptics/woptics_ug_optics.6.29.html)
- BAND: [Band-Gap Analysis of a Photonic Crystal (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.bandgap_photonic_crystal/bandgap_photonic_crystal.html)
- BEND: [Step-Index Fiber Bend (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.step_index_fiber_bend/step_index_fiber_bend.html)
- STRESS: [Stress–Optical Effects in a Photonic Waveguide (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.stress_optical/stress_optical.html)
- GPS: [Stress–Optical Effects with Generalized Plane Strain (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.stress_optical_generalized/stress_optical_generalized.html)
- ANISO: [Optically Anisotropic Waveguide (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.optically_anisotropic_waveguide/optically_anisotropic_waveguide.html)
- MCF: [Directional Coupler (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.directional_coupler/directional_coupler.html)
- SPLICE: [Single Mode Fiber-to-Fiber Coupling (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.single_mode_fiber_coupling/single_mode_fiber_coupling.html)
- FIBER: [Fiber Simulator (6.3, PDF pp.4–5)](https://www.comsol.com/model/download/1423651/applications.fiber_simulator.pdf)
- PCF: [Photonic crystal fiber blog model (Application ID 75371)](https://www.comsol.com/model/photonic-crystal-fiber-blog-model-75371)
- OPT: [The General Optimization Study (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.opt/opt_ug_solver.8.06.html)
- OPTX: [Optimization of a Photonic Crystal for Signal Filtering (6.4)](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.photonic_crystal_filter_optimization/photonic_crystal_filter_optimization.html)
- MAL: [Malitson, Interspecimen Comparison of the Refractive Index of Fused Silica (1965)](https://doi.org/10.1364/JOSA.55.001205)

모든 링크는 2026-11-19 감사에서 열람한 공식 COMSOL 문서/모델 페이지 또는 출판사 원문 메타데이터다. 자료에 없는 물리 상수·제품 수치·문헌 결과는 추가하지 않았다.
