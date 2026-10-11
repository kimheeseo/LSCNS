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

## Phase 7 — 공유 BOM / 호환성 검토

1. **공통 BOM / 호환성**을 열면 현재 `lsdc-twin-bom/1.0` 패킷을 계산합니다.
2. 설정에서 랙 수·GPU 비율을 바꾸거나 광 패널에서 토폴로지·거리·단선 상태를 변경하면 열린 BOM 패널이 약 0.4초 간격으로 갱신됩니다. **현재 설계 다시 계산**으로 수동 갱신할 수도 있습니다.
3. 랙·GPU 서버·GPU·NIC·Leaf/Spine/Core·활성 광 링크·총 길이를 확인합니다. 랙 수 변화가 스위치와 링크를 자동 증설하지는 않습니다. 토폴로지는 공유 패킷의 노드·링크 기준입니다.
4. **검토 JSON**을 저장하거나 **전체 설계 → BOM 도구**로 입력 초안을 전달합니다. 기존 BOM 도구는 별도의 설계 계산 기준을 적용하므로 구매 BOM 수량과 이 요약이 항상 같다고 주장하지 않습니다.

GPU 서버는 명시 서버 수량 또는 랙 수×GPU 서버 비율을 사용합니다. 서버당 GPU/NIC 값을 주면 그 값을 사용하며, 없으면 8 GPU/1 NIC를 가정합니다. 근거에 실제 사용값과 입력/가정을 표시합니다. 필수 입력이 없으면 미확정이며 0대와 구분합니다.

- **PASS**: 제공된 근거가 패킷 규칙 검사를 통과함. 공식 URL 원문·실제 제품의 독립 검증이나 구매 승인은 아닙니다.
- **UNKNOWN**: SKU, 데이터시트 URL, reach, endpoint 프로토콜/포트, 광 connector/FEC 등 필수 근거가 부족함.
- **FAIL**: 프로토콜·속도·reach·매체·connector/FEC 상충, 포트 초과, CPO 중복 또는 잘못된 스키마.

`compatible`은 PASS일 때만 true입니다. 차단 오류가 없다는 의미는 `nonBlocking`을 확인합니다. UNKNOWN을 호환 확정으로 취급하면 안 됩니다. 제품별 세부 광 lane, MPO polarity, 장비 설정 및 공식 지원 목록은 별도 검증 대상입니다.

## Phase 8 — 전체 통합 점검

**종합 점검 / 가이드 → 통합 점검 실행**을 눌러 7개 구역, 랙 자산, Phase 1 FOV, 35개 MEP, 광 경로, Phase 4 GPU 캐시, Phase 5 계측, Phase 6 간섭 계산기, Phase 7 공통 패킷 계산기 등의 API 준비 상태를 검사합니다. 서버실 모델은 모드 전환 없이 읽기 전용으로 검사하며 운영 모드·장애·카메라·선택 설비·설정·로그를 보존합니다.

더 강한 검증은 GitHub Actions의 [Phase 4~8 Chromium 회귀 테스트](https://github.com/kimheeseo/LSCNS/actions/workflows/ls-datacenter-3d-phase8-browser.yml)로 수행합니다. 성공 실행에서만 완료된 테스트 건수와 스크린샷을 검증 결과로 인정합니다. 물리적 Windows GPU/Android/iPhone 성능, 정식 GPU Instancing, 실제 BIM 및 구매 승인은 여전히 별도 과제입니다.

## 소스 구성

동일 폴더에 3D HTML, Extensions JS/CSS, Phase2A/2B JS/CSS 및 아래 5개 추가 파일을 함께 둡니다.

- LS_Datacenter_3D_Phase45.js — GPU 캐시 상태·실기기 FPS 자가진단
- LS_Datacenter_3D_Phase6.js — Data Hall 개념 배관/버스웨이, 간섭 검사
- LS_Datacenter_3D_Phase7.js — 공통 패킷 BOM 요약 및 호환성 검토
- LS_Datacenter_3D_Phase8.js — 종합 기능 점검
- LS_Datacenter_3D_Phase458.css — PC/모바일 UI

연결된 스크립트를 함께 로드하므로 단순 파일 직접 열기보다 GitHub Pages 또는 로컬 HTTP 서버에서 실행합니다. 변경 전 원본 소스와 테스트 보고서는 GitHub 커밋 히스토리에서 확인할 수 있습니다.
