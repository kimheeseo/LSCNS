# LS Datacenter 3D v4.7.0 — Phase 1~8 통합 사용자 가이드

**실행:** [3D 캠퍼스](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?v=4.7.0) · [BOM 도구](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/) · [GitHub 코드](https://github.com/kimheeseo/LSCNS/tree/main/DCI/DataCenter)

**주의:** 교육·설계 검토용 가상 3D 모델입니다. 실측 BIM/CAD, 실제 제조사 승인 BOM, 준공 검증 및 전기·기계 설계 인허가를 대체하지 않습니다.

## Phase 1 — FOV / GPU 상세

캠퍼스 7구역 선택 → GPU 랙 클릭 → 랙 인스펙터 → 서버 트레이 인출 → GPU/패키지 확대. FOV 로그 눈금에서 확대 단계를 확인합니다. µm/nm 형상은 개념도이며 실제 반도체 CAD가 아닙니다.

## Phase 2-A/B — 작업 기록 / 광 연결

- 화면의 **MEP 작업**: 7개 구역별 5개, 총 35개 검토 항목. 체크 상태와 작업 로그는 브라우저에 저장됩니다. CSV 출력 가능.
- **광 연결**: GPU/NIC → ToR/Leaf → Spine A/B → Core. DAC/AEC 구리 구간과 SMF/MPO 광 구간을 구분합니다.
- CPO/Pluggable 방식, 광 단선, B 경로 우회, 정상 복구를 시험합니다. 전력·지연·대역폭 수치는 가정값입니다.

## Phase 3/4 — 렌더 최적화

정적 GPU 지오메트리 캐시, 프러스텀 컬링, 원거리 랙/NPC LOD, 라우트 구간 컬링이 적용돼 있습니다. **FPS / 기기 검증** 패널에서 품질(낮음/보통/높음), 캐시 켜기·끄기를 비교할 수 있습니다.

반복 오브젝트의 정식 GPU Instancing 셰이더는 현재 미구현입니다. 기존 WebGL 불투명·선·투명 지오메트리 배치 방식은 유지했습니다.

## Phase 5 — Windows GPU, Android, iPhone 실제 성능 측정

1. 실제 대상 PC/휴대폰에서 위 3D 페이지를 실행합니다. PC Chrome/Edge, Android Chrome, iOS Safari별로 측정합니다.
2. 동일 구역·카메라·운영 시나리오를 선택하고 **FPS / 기기 검증**을 엽니다.
3. 품질을 지정한 뒤 **10초 성능 측정**을 누르고 탭을 전환하지 않습니다.
4. FPS·p95 프레임 시간·GPU 캐시 업로드 값을 확인하고 **JSON 저장**을 누릅니다.
5. 보통/낮음, 캐시 ON/OFF를 각각 동일 조건에서 반복하고 드라이버, 발열, 배터리 모드 등 장치 정보를 별도 기록합니다.

브라우저 WebGL이 그래픽카드 이름을 비공개할 수도 있습니다. 프레임 간격 기반 지표는 GPU 순수 실행 시간과 다릅니다. CI SwiftShader 헤드리스 측정은 실기기 검증이 아닙니다. 실제 장치별 JSON 확보 전 Phase 5의 **실기기 검증은 미완료**입니다.

## Phase 6 — 랙·냉각·전력 간섭 사전검사

1. **랙 / MEP 검사**를 열어 냉수 Supply/Return 배관, 상부 버스웨이, 랙 안내선 표시를 선택합니다.
2. **Data Hall 이동**을 눌러 추가 3D 설비 개념도를 확인합니다.
3. 가상 Z 오프셋을 조정하고 이격 기준(기본 0.5m)을 입력합니다.
4. **충돌·간섭 검사**에서 검사한 경로/설비 수와 이격 미달 목록을 확인합니다. **검사 CSV**로 저장합니다.

이 검사는 축 정렬 3D AABB 근사와 가상 좌표를 사용합니다. 장비 문 스윙, 실제 배관 열팽창·압력, 케이블 굽힘 반경, 전기 안전 거리, 화재구획, 구조 하중, 내진 등을 검증하지 않으며 비검출이 시공 승인임을 의미하지 않습니다.

## Phase 7 — NVIDIA NIC / CPO / 광모듈 BOM 정합성

1. **NVIDIA / BOM**을 열고 기존 NVIDIA 제품 카탈로그(JSON)가 로드되는지 확인합니다.
2. InfiniBand/Ethernet, CPO/Pluggable, 400/800 Gb/s, 랙 수, 랙당 uplink 수와 가상 링크 거리를 선택합니다.
3. **BOM 감사 실행** 후 논리 링크 수, 스위치 하한, 케이블 수/길이, 부품별 미산정·경고를 확인합니다.
4. **검증 JSON**을 저장하여 기존 BOM 도구의 설계안과 비교합니다. 현재 양방향 수량 동기화는 미구현입니다.

특히 제조사 카탈로그 근거로 다음 원칙을 적용합니다.

- Q3450-LD: InfiniBand CPO 및 MPO12 전면 커넥터. 스위치 **측** 플러거블 광모듈을 추가 계상하지 않습니다. 호스트 모듈 및 실제 송수신 호환성은 미산정입니다.
- Q3400-RA: 논리 포트 144와 물리 OSFP cage 72를 구분합니다. 한 논리 포트에 OSFP 1개로 잘못 산정하지 않습니다.
- ConnectX-8: InfiniBand 800G 단일 링크와 Ethernet 포트 속도는 구분합니다. Ethernet 단일 800G NIC로 확정하지 않습니다.
- Spectrum-4 SN5000: Ethernet 스위치 제품군으로 실제 세부 SKU·포트 수가 없는 경우 구매 수량을 확정하지 않습니다.
- Ethernet 연결에 InfiniBand Q3450-LD CPO를 섞는 등 프로토콜 상충 조건은 차단 경고를 표시합니다.

링크 수=랙 수×uplink/랙, 케이블 길이=링크 수×가상 거리라는 설계 *가정*을 사용합니다. MPO polarity, FEC, host port, optical reach, cooling 및 exact SKU를 공식 문서로 검증해야 실제 BOM이 됩니다.

## Phase 8 — 전체 통합 점검

**종합 점검 / 가이드 → 통합 점검 실행**을 눌러 7개 구역, 랙 자산, Phase 1 FOV, 35개 MEP, 광 경로, Phase 4 GPU 캐시, Phase 5 계측, Phase 6 간섭 계산기, Phase 7 카탈로그 계산기 등의 API 준비 상태를 검사합니다. 이는 시나리오를 강제로 바꾸지 않는 비파괴적 스모크 테스트입니다.

더 강한 검증은 GitHub Actions의 [Phase 4~8 Chromium 회귀 테스트](https://github.com/kimheeseo/LSCNS/actions/workflows/ls-datacenter-3d-phase8-browser.yml)로 수행합니다. 성공 실행에서만 완료된 테스트 건수와 스크린샷을 검증 결과로 인정합니다. 물리적 Windows GPU/Android/iPhone 성능, 정식 GPU Instancing, 실제 BIM 및 구매 승인은 여전히 별도 과제입니다.

## 소스 구성

동일 폴더에 3D HTML, Extensions JS/CSS, Phase2A/2B JS/CSS 및 아래 5개 추가 파일을 함께 둡니다.

- LS_Datacenter_3D_Phase45.js — GPU 캐시 상태·실기기 FPS 자가진단
- LS_Datacenter_3D_Phase6.js — Data Hall 개념 배관/버스웨이, 간섭 검사
- LS_Datacenter_3D_Phase7.js — NVIDIA 제조사 카탈로그 BOM 감사
- LS_Datacenter_3D_Phase8.js — 종합 기능 점검
- LS_Datacenter_3D_Phase458.css — PC/모바일 UI

NVIDIA 공식 카탈로그 파일을 불러오므로 단순 파일 직접 열기보다 GitHub Pages 또는 로컬 HTTP 서버에서 실행합니다. 변경 전 원본 소스와 테스트 보고서는 GitHub 커밋 히스토리에서 확인할 수 있습니다.
