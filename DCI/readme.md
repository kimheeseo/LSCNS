# AI Data Center · DCI & COMSOL PJT

AI 데이터센터 탐색, CPO 공급망 조사, 인프라 설계 및 광학 수치해석 프로젝트 모음입니다.

## 1. 3D Datacenter Tour

- **목적:** 데이터센터 구조와 주요 설비를 3D로 탐색하는 교육·설계 검토용 투어.
- **주요 기능:** 캠퍼스 7개 구역 탐색, 설비 정보 조회, 랙 전·후면 및 내부 점검, 작업자 시점 전환.
- **실행 (Three.js 3D):** [3D Datacenter Tour](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three&v=20261026-display-mode-fix10)
- **참고:** [사용 안내](./DataCenter/LS_Datacenter_3D_사용법.txt) · [2D Campus Tour](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_Campus.html)

교육·검토용 개념 모델이며, 실측 CAD/BIM 또는 실시간 시설 데이터 기반 모델은 아닙니다.

## 2. CPO Supply Chain

- **목적:** CPO(Co-Packaged Optics) 구성과 부품별 공급망 파악.
- **주요 기능:** 국가·부품별 업체 및 제품 조회, 국내 영업·문의 채널 확인, CPO 조립 구조와 광 경로 2D/3D 탐색, 출처 기반 시장 동향·행사 일정 확인.
- **실행:** [CPO Supply Chain Explorer](https://kimheeseo.github.io/cpo-supply-chain/)
- **자료:** [별도 저장소](https://github.com/kimheeseo/cpo-supply-chain)

제품·기술의 발표, 시연, 양산 단계는 각 자료의 출처와 상태를 확인합니다.

## 3. Datacenter Tool

- **목적:** AI 데이터센터 인프라 구성 검토 및 BOM 산정.
- **주요 기능:** Compute, Network, Optical, Rack, Power, Cooling 설계·용량 검토, 제품 매핑, 2D/3D 시각화, Excel·CSV·Design JSON 내보내기.
- **실행:** [Datacenter Tool](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/)
- **자료:** [소스·상세 안내](./DataCenter/) · [설계 계산 백서](./DataCenter/whitepaper/v1.4/README.md)

핵심 계산·BOM 제품 매칭 기능은 private backend 연결 상태에 따라 사용 가능 여부가 달라집니다.

## 4. COMSOL PJT · Python Wave Optics

- **목적:** 광도파관·광섬유의 전파 특성과 FEM 수치해석 연구.
- **주요 범위:** M1~M8 도파관·모드 해석, PML·산란 S-파라미터, 분산·코팅 설계, SMF·2코어 MCF·이상적 HCF 모델.
- **주요 결과:** 유효굴절률(`n_eff`), 모드 분포, MFD, 유효면적(`Aeff`), 분산, 결합 길이 및 메시·영역 수렴 검토.
- **실행:** [Wave Optics Lab](https://kimheeseo.github.io/LSCNS/COMSOL/) · [Workbench 미리보기](https://kimheeseo.github.io/LSCNS/COMSOL/workbench/)
- **자료:** [프로젝트 안내](../COMSOL/README.md) · [설치·사용 가이드](../COMSOL/workbench/USER_GUIDE.md) · [광섬유 계산 예제·Colab 안내](../COMSOL/fiber_simulation_guide/readme.md) · [공개 자료 비교 결과](../COMSOL/public_reference_benchmarks/results/COMPARISON_REPORT.md)

Python 기반 연구 프로젝트입니다. 웹은 해석식·저장 결과 보간 또는 미리보기를 제공하며, 새 FEM 해석은 로컬 Python·Colab에서 수행합니다. COMSOL 프로그램 자체를 실행한 검증이나 실제 제품 적합성 검증을 의미하지 않으며, HCF는 이상적 모세관 모델 범위입니다.
