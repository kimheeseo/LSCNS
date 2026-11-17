# LS Wave Optics — STEP 0 COMSOL 광섬유 기능 감사

감사일: 2026-11-18 (KST) · 대상: kimheeseo/LSCNS, COMSOL/universal_silica_v01 · UI/solver 버전: v0.7.1

## 1. 감사 결론과 범위

**다음 단계 구현은 시작하지 않았다. 애플리케이션 코드 수정, GitHub 커밋, 배포 변경은 없다.** 이 문서는 공식 문서와 현재 소스의 읽기 전용 감사 결과 및 승인용 제안이다.

사용자가 설명한 자산 목록보다 현재 소스가 앞서 있다. 이미 브라우저용 Yee 벡터 솔버, 복소 PML, 임의 단면, 직접 N코어 해석, 스윕/프로젝트 기능이 존재한다. 핵심 과제는 이 기능을 다시 만드는 것이 아니라 검증 수준을 높이고 누락된 물리를 추가하는 것이다. 특히 neff 실수부가 기준값과 가까워도 누설손실 정확도가 확보되었다고 판단할 수 없다.

- 고정 소스: [main snapshot e62a53babf1b](https://github.com/kimheeseo/LSCNS/tree/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01); 저장소 전체 HEAD 시각 2026-11-18 00:54:32 KST. 이 시각은 이 하위 프로젝트의 최종 수정 시각을 뜻하지 않는다.
- 공식 근거: COMSOL 6.4 Wave Optics User’s Guide와 Application Library 중심. Fiber Simulator PDF는 확인 가능한 6.3 문서를 별도로 표시했다.
- 확인 방법: 대상 하위 트리·관련 구현·검증 문서의 정적 검토. 이번 감사에서 COMSOL .mph 실행, 브라우저/Node 재벤치마크 또는 새로운 수치 시험은 하지 않았다.
- 아래 'COMSOL 구현 방식'은 문서로 확인한 사실이다. '난이도·필요 알고리즘·검증 방법·우선순위'는 우리 툴에 대한 엔지니어링 제안이며 COMSOL 공식 보증이나 요구사항이 아니다.
- '미확인'은 조사한 공식 문서에서 해당 세부 기능을 확인하지 못했다는 뜻이다. '미검증'은 코드나 산출값이 있어도 정확도 검증이 부족하다는 뜻이다.

## 2. 저장소 구조와 파일 맵

상위 COMSOL/에는 universal_silica_v01 외에 COMSOL_Install, research, public_reference_benchmarks, g657a1_validation, validation_report와 검증 실행 스크립트가 있다. 상세 감사 범위는 사용자가 지정한 universal_silica_v01 및 그 웹 진입점이다. 다른 연구 폴더 전체에 대한 부재 판정은 하지 않았다.

| 계층 | 주요 파일 | 역할·감사 결과 |
| --- | --- | --- |
| 웹 진입·기존 계산 | [web/index.html](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/index.html), [web/app.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/app.mjs), [web/solver.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/solver.mjs) | v0.7.1; radial P1 scalar LP01, 재료·분산·메시·결과 패널의 기존 경로 |
| 재료 | [web/material_models.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/material_models.mjs), [web/m2_preview_engine_v06.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/m2_preview_engine_v06.mjs), [web/m2_preview_ui_v06.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/m2_preview_ui_v06.mjs) | 도핑/표 데이터, 보간, n/k/α; 모델과 계수 출처·유효 범위 검증은 계속 필요 |
| 기존 벡터 경로 | [web/vector_fdfd_browser.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/vector_fdfd_browser.mjs), [web/vector_m3_ui.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/vector_m3_ui.mjs), [web/m3_study_ui_v07.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/m3_study_ui_v07.mjs) | collocated Hx/Hy 벡터 FDFD. 새 Yee 경로와 다른 이산화이며 비교용으로 보존 |
| 현재 Yee 계산 코어 | [web/yee_solver.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/yee_solver.mjs), [web/yee_sparse.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/yee_sparse.mjs), [web/vector_study_worker.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/vector_study_worker.mjs) | staggered transverse-E 벡터 고유계, 복소 CSR/ILU-GMRES; Worker 실행 |
| 공통 단면 | [web/vector_geometry.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/vector_geometry.mjs) | 동심원·polygon·mask·HCF·MCF를 ε 격자로; subpixel/PML/등가 굽힘 |
| 새 작업 공간 | [web/vector_workbench.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/vector_workbench.mjs), [web/vector_study.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/vector_study.mjs), [web/m3_yee_quickstart_v071.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/m3_yee_quickstart_v071.mjs) | M3–M6 진입, 복수 모드, overlay, JSON schema 3, 1–2축 sweep; v0.7.1은 기존 Yee 노출 개선 |
| HCF/MCF/PMD 기존 자산 | [web/hcf_designer.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/hcf_designer.mjs), [web/mcf_xt.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/mcf_xt.mjs), [web/pmd.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/pmd.mjs) | ARROW/NANF 단면, 2코어 참조점 보간, 통계 PMD 시연. 직접 Maxwell/응력광학과 구분 |
| Python radial | [universal_radial.py](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/universal_radial.py), [Universal_Silica_Fiber_v0_1.ipynb](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/Universal_Silica_Fiber_v0_1.ipynb) | scalar LP01 기준 경로 |
| Python full-vector | [full_vector/](https://github.com/kimheeseo/LSCNS/tree/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/full_vector), [full_vector/vector_modes.py](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/full_vector/vector_modes.py) | SciPy collocated Hx/Hy FDFD, real isotropic, PML 없음; 원형 HE11 정확해 함수 보유 |
| Python HCF | [hcf_complex_fem/](https://github.com/kimheeseo/LSCNS/tree/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/hcf_complex_fem), [hcf_complex_fem/hcf_complex_fem.py](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/hcf_complex_fem/hcf_complex_fem.py) | complex PML을 쓰는 삼각형 P1 **scalar** FEM; full-vector 비교 기준으로 동일시 금지 |
| Python 굽힘 | [bending_fem/](https://github.com/kimheeseo/LSCNS/tree/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/bending_fem), [bending_fem/solver.py](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/bending_fem/solver.py) | scalar 2D P1 + 등가 굽힘 + PML. 파일 존재와 현재 UI 연결 여부는 별개 |
| 전파·3D 표시 | [propagation_3d/](https://github.com/kimheeseo/LSCNS/tree/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/propagation_3d), [propagation_3d/propagation_3d.py](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/propagation_3d/propagation_3d.py), [web/fiber_3d_core.mjs](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/fiber_3d_core.mjs) | 2+1D scalar split-step FFT BPM과 3D 시각화. 3D full-vector Maxwell 해석이 아님 |
| 검증 근거 | [web/M2_M3_V071_ACCEPTANCE.md](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/M2_M3_V071_ACCEPTANCE.md), [web/M3_M6_YEE_VALIDATION_20261112.md](https://github.com/kimheeseo/LSCNS/blob/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/M3_M6_YEE_VALIDATION_20261112.md), [web/tests/](https://github.com/kimheeseo/LSCNS/tree/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/tests), [web/results/yee_validation/](https://github.com/kimheeseo/LSCNS/tree/e62a53babf1bcbfe5cb66d311ce1a82f7c63505f/COMSOL/universal_silica_v01/web/results/yee_validation) | 구현 여부와 수용 기준 미통과 항목을 분리해 기록 |

index.html은 기존 scalar·Colab import·collocated vector 패널과 vector_workbench를 함께 로드한다. 후속 변경은 이 공존 구조와 JSON 호환 경로를 보존해야 한다.

## 3. 솔버 현황: 구현과 검증을 분리

| 기능 | 현재 코드 | 연구용 판정 |
| --- | --- | --- |
| scalar radial LP01/분산 | 웹·Python 존재 | 기존 기준 경로로 보존. 약유도·축대칭 범위를 벗어나 일반 벡터 해석으로 간주하지 않음 |
| 브라우저 일반 2D 벡터 | Yee-FDFD 1.0 experimental 존재 | step-index 마지막 두 격자 neff 수렴 기준 미통과 |
| PML/복소 neff/재료흡수 | Yee에 연결되어 있음 | 거의 무손실 케이스의 Im neff·손실 수렴 미통과 |
| HCF/NANF 복소 모드 | 기존 단면을 Yee/PML로 샘플링 | 얇은 벽 under-resolution으로 손실 정량 미검증 |
| 임의 단면·mask | 영역·색상 mask, SVG 래스터화, 4×4 ε averaging 존재 | 작은 구조·고대비 경계에 대한 subpixel 정확도 검증 필요; 벡터 SVG conforming mesh와 다름 |
| 굽힘 | n·exp(x/R) 등방 등가 굴절률 | 완전한 곡선 좌표 벡터식 또는 공식 축대칭 모델과 동등성 미검증 |
| 직접 N코어 | supermode, 2코어 splitting, N코어 reduced coupling 존재 | 근사 투영/코히런트 결합 모델. 실제 random-coupling cable XT와 구분 |
| workflow/sweep/project | Worker·취소·36점 sweep·이력·schema3 존재 | 과거 브라우저 검증 기록 있음. mode-index 정렬은 mode-following 인증이 아님 |
| Floquet/Bloch ω 고유계 | 대상 계산 코어에서 확인되지 않음 | 새 물리·연구 유형 필요 |
| 재료 ε tensor/응력광학 | 대상 일반 Yee는 isotropic sampling | 응력 해석·광탄성 tensor 결합·일반 tensor Maxwell 확장 필요 |
| overlap 접속/반사 손실 | 완성된 일반 E/H port/접속 해석 경로 확인되지 않음 | mode overlap부터 추가하고 실제 반사·gap은 별도 범위로 명시 |
| 최적화 | sweep는 존재; general optimizer 미확인 | 신뢰 가능한 목적함수 확보 후 추가 |

Yee 알고리즘은 block shift-invert subspace iteration + complex ILU(0)-GMRES + Rayleigh–Ritz이다. 기본 외측 잔차 허용값은 무손실 10⁻⁸, PML/흡수 10⁻¹⁰; 내측은 min(2×10⁻¹⁰, 외측×0.02), GMRES restart 36, 최대 외측 120회/선형 반복 1400회다. 이는 선형대수 잔차 설정이며 공간 이산화 오차나 극소 손실의 정확도 보증이 아니다.

상한은 축당 24–192셀, 요청 1–16모드, sweep 전체 36점, mask 512픽셀이다. 임의 SEM/PCF 구조가 들어간다는 사실만으로 얇은 실리카 벽을 충분히 해상할 수 있다는 뜻은 아니다.

기존 가짜 모드 후보 제외 기준은 PML의 |E|² 비율>35%, 물리 영역<65%, Im neff<−10⁻⁸, 잔차 기준 초과다. 이 비율은 전력 또는 에너지 정규화에 기초한 완전한 물리성 증명이 아니다. 물리 영역 필드와 계산창·PML 변화에 대한 안정성 검사가 추가로 필요하다.

### 이미 저장된 수치 증거 — 이번 감사의 재측정값 아님

출처: 위 M3–M6 검증 문서, 2026-11-12. 환경 Node v24.22.1 / AMD EPYC 9V74 / Linux. 브라우저 사용자 PC의 속도 보장이 아니다.

| 항목 | 저장된 값 | 감사 판정 |
| --- | --- | --- |
| scalar LP01 해석해 / scalar FEM | 1.44610478058 / 1.44610475118 | 약유도 비교 기준이며 정확한 벡터 HE11 해와 다름 |
| Yee N=72 / 96 | 1.44609935051 / 1.44610489953 | 인접 차이 5.549×10⁻⁶ > 기존 10⁻⁶ 기준: 미통과 |
| N96 vs scalar LP01 해석해 | 차이 1.189×10⁻⁷ | 단일 격자의 작은 오차로 비단조 격자 수렴을 대체할 수 없음 |
| 동일 Yee 행렬 JS / SciPy ARPACK | neff 차이 약 9.5×10⁻¹³ | 고유값 알고리즘 교차 확인; 독립 물리 모델 검증 아님 |
| 직선 PML 변화 | −7.743~14.963 dB/km | 부호까지 불안정: 작은 누설손실 미검증 |
| HCF 얇은 벽 | 0.45 µm 벽, 격자 간격 1.5–2.25 µm | subpixel로 손실 정확도를 보증할 수 없음 |
| N96 실행시간 / Node RSS | 20.955 s / 168.145 MiB | 전체 Node 프로세스 RSS. Worker 전용 peak memory 아님 |
| N96 배열 모델 | 22.782 MiB | 실측 메모리가 아닌 추정 |

## 4. COMSOL 광섬유 연구 기능 매트릭스

약어: EWFD=Electromagnetic Waves, Frequency Domain; EWBE=Electromagnetic Waves, Beam Envelopes. P0=검증 기반 우선, P1=핵심 광섬유 연구 확장, P2=다중물리/최적화 후속. P0/P1은 후보 필터 검증과 분류 확장을 나눠 수행한다는 뜻이다.

| 기능 | COMSOL 구현 방식: 인터페이스·경계조건·연구 | 광섬유 연구 용도 | 우리 툴 난이도 | 필요 알고리즘: 제안 | 검증 방법: 제안 | 우선순위 | 공식 근거: 제목/URL |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 2D Full-vector 모드·복소 neff | EWFD, 2D Mode Analysis; 전파상수 고유값, 개방계에는 PML/SBC | PCF/HCF/MCF의 Ex/Ey/Ez, β, 복수 유도·누설 모드 | 상 | 기존 Yee 복소 희소 고유계 재사용; shift-invert; 독립 벡터 기준해 | 원형 HE11 정확해 + 약유도 LP01 구분, mesh/창 크기 수렴, 독립 이산화와 비교 | P0 | [Leaky Modes in a Microstructured Optical Fiber](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html) |
| PML·누설손실·재료흡수 | EWFD + PML 도메인 + Mode Analysis; 복소 neff | 유한 클래딩 confinement loss와 재료흡수 | 상 | 기존 complex stretching 교정; Im(neff) 정밀도, 흡수 연결, 가짜 모드 판별 | 격자·물리영역·PML 두께/강도/차수 독립 변화; κ=0/비영 비교; 공식 MOF 표 | P0 | [Leaky Modes in a Microstructured Optical Fiber](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html) |
| Scattering BC | EWFD Scattering Boundary Condition; mode용 propagation-constant subtraction | 외부 방사 경계 및 PML 결과 교차 확인 | 중~상 | 벡터 방사 경계; β 의존 경계항 반복/비선형 고유계 처리 | 동일 MOF의 PML/SBC 복소 neff 및 경계 반사·영역 수렴 | P1 | [Leaky Modes in a Microstructured Optical Fiber](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html); [Wave Optics Module User’s Guide](https://doc.comsol.com/6.4/doc/com.comsol.help.woptics/WaveOpticsModuleUsersGuide.pdf) pp.162–165 |
| 유한 PCF 클래딩·결함 코어 설계 | 공기공 배열 Geometry/Materials + EWFD Mode Analysis; 유한 외부 경계 | 공기공 d/Λ, 링 수, 결함, index guiding과 leakage 평가 | 중 (솔버 검증 포함 상) | 기존 영역/마스크 격자 재사용; 삼각/육각 격자와 결함 생성, subpixel 영향 | 공식 PCF mode-dispersion 예제; 링 수·경계·해상도 변화 및 독립 mesh 비교 | P1 | [Photonic crystal fiber blog model (Application ID 75371)](https://www.comsol.com/model/photonic-crystal-fiber-blog-model-75371) |
| 고유주파수 밴드 구조·Floquet | EWFD + unit-cell Floquet periodic BC + Eigenfrequency + k sweep | 무한 주기 클래딩의 Bloch band/gap | 상 | 복소 Bloch 위상 경계, k 경로, ω 고유계; 분산 재료는 비선형 고유계 | 균일 매질의 plane-wave dispersion + 공식 band diagram; 필요시 Brillouin zone 내부도 검사 | P1 | [Band-Gap Analysis of a Photonic Crystal](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.bandgap_photonic_crystal/bandgap_photonic_crystal.html) |
| 굽힘: 등각 변환/축대칭 | 공식 fiber bend는 2D Axisymmetric Mode Analysis + PML; 방위각 전파. 등각 exp(x/R) 내장 기능은 미확인 | macro-bend mode·loss와 반경 의존성 | 상 | 현재 등가 굴절률 근사 범위 검증; 곡선 좌표 벡터식 또는 축대칭 백엔드 | 공식 Step-Index Fiber Bend, R→∞ 직선 극한, 동일 모드 R sweep와 PML 수렴 | P1 | [Step-Index Fiber Bend](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.step_index_fiber_bend/step_index_fiber_bend.html) |
| 분산·파장 스윕 | Mode Analysis + wavelength Parametric Sweep; Fiber Simulator 결과 | ng, D, S0, ZDW, 모드별 분산 | 중 | 기존 재료 n(λ) + 벡터 neff; 필드 overlap 모드 추적, 미분 간격/차수 검증 | 재료 분산 분리; λ 간격 수렴; scalar 약유도 극한과 공식 Fiber Simulator | P1 | [Fiber Simulator (Application Library, 6.3)](https://www.comsol.com/model/download/1423651/applications.fiber_simulator.pdf); [Optically Anisotropic Waveguide](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.optically_anisotropic_waveguide/optically_anisotropic_waveguide.html) |
| 복굴절·편광·모드 추적 | EWFD tensor permittivity + Mode Analysis; mode following | 형상/재료 복굴절, 편광축, beat length | 상 (형상만 중) | 형상 비대칭은 기존 등방 Yee 검증; 재료 비등방은 전체 ε tensor 확장; 퇴화 부분공간 추적 | 원형 대칭의 퇴화, 회전 불변성; 공식 anisotropic 예제 비교 | P1 | [Optically Anisotropic Waveguide](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.optically_anisotropic_waveguide/optically_anisotropic_waveguide.html) |
| 응력광학 결합 | Solid Mechanics Stationary → photoelastic 굴절률 변화 → EWFD Mode Analysis; Structural Mechanics 필요 | 잔류·열·인장 응력의 복굴절, PM fiber | 상 | 2D 탄성/열팽창 FEM + 광탄성 tensor + 광학 격자 매핑; 계수 출처 관리 | 균일 응력 해석해; 응력·광학 mesh 별도 수렴; 공식 stress-optical 예제 | P2 | [Stress–Optical Effects in a Photonic Waveguide](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.stress_optical/stress_optical.html) |
| MCF supermode·coupling | 전 단면 Mode Analysis; Directional Coupler는 수치 port/Boundary Mode Analysis와 EWBE 전파 예제 | 2/N코어 supermode splitting, beating, 코어 간 결합 | 중~상 | 기존 직접 N코어 경로 검증; 전력 정규화·퇴화 처리·결합행렬; random XT는 별도 전파 통계 | 동일 2코어 even/odd Δβ, Lc, power conservation; 큰 pitch 결합 소멸; N코어 대칭 | P1 | [Directional Coupler](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.directional_coupler/directional_coupler.html); N코어 광섬유 전용 공식 예제: 미확인 |
| Overlap 접속·결합 손실 | Numeric Port + Boundary Mode Analysis, EWBE Frequency Domain; 공식 예제는 렌즈형 fiber-to-fiber 결합 | 횡오정렬·회전·MFD 불일치 coupling; 단순 접속 근사 | 중 (반사·gap 포함 상) | E/H 전력 overlap·위상 정렬; 반사/축방향 불연속은 mode matching 또는 별도 전파 해석 | 동일 모드 η=1; Gaussian 이동 해석해; 좌우 교환/대칭; port S-parameter 비교 | P1 | [Single Mode Fiber-to-Fiber Coupling](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.single_mode_fiber_coupling/single_mode_fiber_coupling.html); [Wave Optics Module User’s Guide](https://doc.comsol.com/6.4/doc/com.comsol.help.woptics/WaveOpticsModuleUsersGuide.pdf) pp.82–87 |
| 파라메트릭 스윕·저장·비교 | Parametric Sweep + Mode Analysis/Eigenfrequency; mode following | a, d/Λ, λ, R, 도핑의 설계 지도 | 하~중 | 기존 Worker queue/취소/JSON/CSV 보존; sweep마다 모드 추적·수렴 상태 기록 | 단일 실행 vs 동일 sweep 점; 저장/이관 roundtrip; 실패·취소 결과 분리 | P1 | [Fiber Simulator (Application Library, 6.3)](https://www.comsol.com/model/download/1423651/applications.fiber_simulator.pdf); [Optically Anisotropic Waveguide](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.optically_anisotropic_waveguide/optically_anisotropic_waveguide.html) |
| 최적화 | General Optimization study; Optimization Module 필요 | loss/Aeff/D/복굴절 등 목적·제약 최적화 | 상 | 검증된 sweep/cache 후 bounded derivative-free 최적화; 필요시 고유값 감도/adjoint | 단순 기준 문제 + 전체 grid sweep 대비; 경계해·재시작; 최종 독립 재계산 | P2 | [The General Optimization Study](https://doc.comsol.com/6.4/doc/com.comsol.help.opt/opt_ug_solver.8.06.html) |
| 후처리: MFD·Aeff·γ | Fiber Simulator의 MFD/field plot 확인. Aeff·γ 전용 내장 자동 기능 및 해당 MFD 정의: 미확인 | 결합 효율, 비선형성, 모드 크기 | 중 (일반 벡터 γ 상) | MFD 정의 명시; 강도 모멘트·적분; 균질 약유도 γ 근사와 재료별 vector overlap 구분 | Gaussian 정확해; 크기 scaling; 재료 n2와 적분 영역/정규화 검사 | P1 | [Fiber Simulator (Application Library, 6.3)](https://www.comsol.com/model/download/1423651/applications.fiber_simulator.pdf); Aeff·γ 자동 기능: 미확인 |
| 후처리: 모드 분류·PML 가짜 모드 | MOF 예제 HE/TE/TM-like 분류; anisotropic 예제 mode following. 범용 가짜 모드 자동 판정 기준은 미확인 | 물리 모드 선별·고차 모드 추적 | 중~상 | 편광률·대칭·절점·E/H 구조; PML participation + domain/PML 안정성; 퇴화군 처리 | 대칭 원형 HE11/TE01/TM01 기준; 격자/PML 변경에도 필드/고유값 추적 | P0/P1 | [Leaky Modes in a Microstructured Optical Fiber](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html); [Optically Anisotropic Waveguide](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.optically_anisotropic_waveguide/optically_anisotropic_waveguide.html) |

## 5. 감사에서 반드시 구분해야 할 물리

### 5.1 유한 PCF 모드와 무한 주기 클래딩 bandgap

유한 단면은 λ를 고정하고 β 또는 neff를 구한다. Bloch band는 주기 셀의 k를 주고 ω를 구한다. 현재 transverse-β 고유계에 Floquet 경계만 덧붙여 공식 Eigenfrequency와 같다고 표현하면 안 된다.

공식 Band-Gap Analysis 예제는 2D 주기 포토닉 결정의 근거다. 광섬유의 축방향 βz를 매개변수로 한 projected band diagram에 대한 전용 공식 PCF 절차까지 확인된 것은 아니다. 이 세부 항목은 **미확인**이며, 후속 설계에서 전파 방향·편광·k 경로와 주파수 문제를 먼저 정의해야 한다. 특정 고대칭 경로의 gap만으로 완전한 bandgap을 선언하지 않는다.

공식 MOF 누설 예제는 high-index-core 구조다. 이것을 hollow-core bandgap 또는 NANF antiresonant guiding의 직접 검증으로 대체하지 않는다. PCF, PBGF, AR-HCF/NANF는 구조와 유도 메커니즘을 구분해 표시해야 한다.

### 5.2 복소 neff 부호·손실 정의

현재 코드: exp(iβz−iωt), ñ=n+iκ, β=k0·neff, k0=2π/λ. 따라서 수동 매질에서 Im(neff)>0가 감쇠다.

COMSOL User’s Guide: exp(iωt−αz), α=δz+iβ, neff=−iα/k0. 따라서 같은 수동 전진 모드의 neff 허수부 부호가 현재 코드와 반대다. 공식 MOF 표의 HE11 PML 값도 1.44539464−3.18614991×10⁻⁸i이다. 비교는 시간·전파 규약을 변환한 뒤 수행한다. 절댓값을 취해 음수 손실을 숨기지 않는다.

현재 규약에서의 전력 손실 환산은 다음과 같다. λ는 m 단위다.

$$\alpha_{\mathrm{dB/km}}=\frac{20}{\ln 10}\,\frac{2\pi}{\lambda}\,\mathrm{Im}(n_{\mathrm{eff}})\times1000.$$

이는 진폭이 exp(−k0 Im(neff) z)로, 전력이 그 제곱으로 감소하는 데서 나온다. COMSOL 규약의 neff 허수부를 그대로 사용할 때에는 위 식 앞의 부호가 바뀐다. αtotal에는 재료흡수와 방사 누설이 함께 들어간다. κ=0의 방사손실, 흡수를 포함한 총손실, 직선/굽힘 차이는 같은 모드를 추적한 별도 계산으로 보고한다. 모드가 재분포하므로 단순 차이가 보편적인 독립 손실 분해라고 보증하지 않는다.

### 5.3 굽힘·MCF·접속의 범위

- COMSOL 공식 굽힘은 굽힘 중심 주위 방위각 전파를 다루는 2D axisymmetric 모델이다. 직선 섬유 단면 자체가 동심원이어야 한다는 뜻으로 오해하면 안 된다. 현재 등가 index 근사와 동일 모델이라는 근거는 미확인이다.
- 2개 동일 코어의 even/odd supermode는 κ=|βeven−βodd|/2, 완전 전달 길이 Lc=π/(2κ)로 연결된다. 굽힘/응력/길이 방향 불균일을 포함한 실제 MCF 누화는 추가 모델이 필요하다.
- 단순 overlap은 두 모드의 정렬·크기 불일치 평가에 쓴다. Fresnel 반사, 공기 gap, 테이퍼, 산란이 있는 접속부의 완전한 삽입손실은 별도 mode matching/전파 문제다. 공식 Single Mode Fiber-to-Fiber Coupling은 렌즈 결합 예제로, 용융접속 자체의 검증 예제는 아니다.

### 5.4 MFD·Aeff·γ와 모드 명명

공식 Fiber Simulator에서 MFD 출력은 확인했다. 그 PDF만으로 MFD가 Petermann/second-moment/1/e² 중 어느 정의인지 확인하지 못했다. Aeff·γ의 전용 자동 변수도 이번 조사에서는 미확인이다. 이 항목을 'COMSOL 기본 버튼과 동일'로 표시하지 않는다.

우리 후처리는 정의를 산출값과 함께 저장해야 한다. 약유도 균질 매질의 통상적인 강도 적분 Aeff=(∫I dA)²/(∫I² dA), γ≈2πn2/(λAeff)는 먼저 범위를 한정해 사용할 수 있다. PCF/HCF의 공간별 n2·고대비·일반 벡터 모드에는 재료별 비선형 overlap과 적절한 전력 정규화가 필요하며 이 단순식을 보편식으로 쓰지 않는다. 비원형 모드는 x/y MFD를 함께 제공하는 방안을 제안한다.

HE/TE/TM/LP 이름은 대칭과 근사 조건을 확인해 붙인다. 임의 PCF에 Ex 우세라는 이유만으로 정확한 HE11 이름을 확정하지 않는다. 퇴화 쌍은 개별 벡터의 임의 회전에 민감하므로 부분공간을 추적해야 한다.

## 6. 미확인·미검증 항목 등록부

| 항목 | 상태 | 해소 방법 |
| --- | --- | --- |
| COMSOL 내장 exp(x/R) 등각 굽힘 옵션 | 미확인 | 특정 기능 노드·수식 문서 확보; 그 전에는 우리 근사로 표시 |
| N코어 광섬유 전용 공식 Application Library 모델 | 미확인 | 공식 directional coupler를 일반 원리 근거로만 사용; 별도 MCF 기준 확보 |
| PCF βz-projected band 전용 공식 절차 | 미확인 | 벡터 unit-cell 문제의 축/고유변수를 명시하고 외부 기준 모델 확보 |
| COMSOL MFD 세부 정의, Aeff·γ 자동 내장 출력 | 미확인 | 앱 내 식/모델 변수 문서를 확인한 뒤 동일성 판단 |
| 보편적인 COMSOL PML 가짜 모드 자동 필터 | 미확인 | 필드/PML/domain 안정성 기준을 우리 검증 규칙으로 명시 |
| 우리 Yee 공간 수렴·초저 누설손실 | 미검증/기준 미통과 | P0에서 mesh·영역·PML 각각의 수렴과 독립 방법 검증 |
| 기존 HCF Colab vs Yee | 같은 벡터 기준 아님 | scalar/full-vector, 입력·모드·규약 차이를 비교표에 명시 |
| 실제 도핑별 광탄성·열·기계 계수 | 이번 감사에서 미검증 | 논문/공식 물성 출처와 농도·온도 범위를 확보 |
| 최신 코드의 재실행 정확도·사용자 브라우저 속도 | 이번 감사 미측정 | 승인 후 지정 benchmark와 환경을 고정하여 재측정 |

## 7. 승인 대상: 추가 개발 순서 제안

1. **P0 — 현재 구현을 보존하면서 기준 확정.** 기존 scalar·Colab 경로·JSON roundtrip 회귀 기준을 고정한다. 현재 Yee를 재사용하며 원형 정확 벡터해, 격자/창 수렴, PML loss convergence, 복소 부호·정규화·모드 추적을 검증한다. 전체 ε tensor 지원 여부는 별도 기능으로 관리한다.
2. **P1 — 유한 PCF 연구 경로.** 기존 geometry adapter에 격자공/결함 파라미터를 추가하고, 검증된 vector 모드에서 분산·MFD/Aeff·birefringence·MCF supermode·overlap을 연결한다. 현재 worker/UI/import 경로를 재작성하지 않는다.
3. **P1 — 별도 Bloch 연구 유형과 개방 경계/굽힘 확장.** 유한-fiber Mode Analysis와 주기-cladding Eigenfrequency를 독립 study로 제공한다. SBC 및 곡선 좌표/축대칭 굽힘은 물리 동등성을 검증한다. 모든 밴드·모드 출력에 study 종류를 기록한다.
4. **P2 — 응력광학과 최적화.** 기계·열·광탄성 입력 및 tensor Maxwell 결합을 추가한다. 최적화는 수렴한 목적함수와 모드 식별이 확보된 뒤 연결한다.

이 순서는 STEP 0에 대한 승인용 제안이다. 후속 단계의 구현·커밋·배포를 수행하지 않았다. 전체 COMSOL 연구 범위의 동등성이나 수치 정확도를 이 감사만으로 선언하지 않는다.

## 8. 공식 문서 목록

- **UG**: [Wave Optics Module User’s Guide](https://doc.comsol.com/6.4/doc/com.comsol.help.woptics/WaveOpticsModuleUsersGuide.pdf)
- **MOF**: [Leaky Modes in a Microstructured Optical Fiber](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html)
- **BEND**: [Step-Index Fiber Bend](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.step_index_fiber_bend/step_index_fiber_bend.html)
- **BAND**: [Band-Gap Analysis of a Photonic Crystal](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.bandgap_photonic_crystal/bandgap_photonic_crystal.html)
- **STRESS**: [Stress–Optical Effects in a Photonic Waveguide](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.stress_optical/stress_optical.html)
- **ANISO**: [Optically Anisotropic Waveguide](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.optically_anisotropic_waveguide/optically_anisotropic_waveguide.html)
- **COUPLER**: [Directional Coupler](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.directional_coupler/directional_coupler.html)
- **SPLICE**: [Single Mode Fiber-to-Fiber Coupling](https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.single_mode_fiber_coupling/single_mode_fiber_coupling.html)
- **FIBER**: [Fiber Simulator (Application Library, 6.3)](https://www.comsol.com/model/download/1423651/applications.fiber_simulator.pdf)
- **PCF**: [Photonic crystal fiber blog model (Application ID 75371)](https://www.comsol.com/model/photonic-crystal-fiber-blog-model-75371)
- **OPT**: [The General Optimization Study](https://doc.comsol.com/6.4/doc/com.comsol.help.opt/opt_ug_solver.8.06.html)

자료 확인 기준일: 2026-11-18 KST. COMSOL 라이선스별 전체 제품 조합을 감사한 문서는 아니다. 응력광학 예제의 Structural Mechanics 및 최적화의 Optimization Module 추가 의존성은 해당 문서에서 확인했다.
