# Gold Pixel Tour · Pixel OJT · 2026-10-13

기준 커밋: `bbd3f4c74ac3364358b9a9343a247bf819606fbd`.

## 변경 전 구조

7개 구역, 42개 정보 항목, `maps[zone] = {grid, objects, doors}`, 16px 타일/384×320 Canvas, `info`·`diagramData`·`dedicatedViews`를 사용했습니다. 조사 대상은 `target()`으로 찾고 `act()`에서 열며, `ls-pixel-discoveries`에 발견 ID를 저장했습니다. 직전 온보딩은 별도 CSS/JS와 상단 학습 화면을 사용했습니다.

## 변경 사항 (diff 요약)

```diff
- Gray Zone: Carrier/MMR/ODF와 UPS/PDU/CDU 혼재
+ Gray Zone: 전기·기계 지원 / Switchgear·ATS·UPS·배터리·분전·CDU·정비
- White Zone: MMR/ODF와 IT 공간 혼재
+ White Zone: IT 랙·Containment·Busway·Rack PDU·Leaf/ToR
- Network/MMR: 통신 인입과 Leaf 랙 중복
+ Network/MMR: Entrance Room·Carrier·MMR·ODF·Core/Spine
+ 구역별 역할 / 들어오는 것 / 나가는 것 3줄 패널
- 별도 상단 온보딩 화면과 CSS/JS 파일 참조
+ 기존 픽셀 지도에 학습 콘솔 6개 / 조사(A)·노트 재열기
- 42개 탐험 정보
+ 58개: 학습 6, 인입/RPP/Entrance Room/칠러B/예비칠러/매니폴드/QD/ColdPlate/누수센서/배터리
+ 캠퍼스 → Utility → Gray → White 전력 동선 / Busway 이후 PDU A/B 분기 / 노란 점선 / 현재 설비 강조
+ 상류·하류·이중화 설명 / 단일·N+1·2N 픽셀 비교 카드
+ 정전: UPS 즉시 보완 → 발전기 정격 확보 → ATS 전환 → 정상 입력 확인
+ CW·FWS·TCS 3루프 / 공급·환수 픽셀 흐름 / 열교환 경계 설명
+ 공랭·D2C·RDHx 픽셀 비교 / 랙 액냉 보기 탭
+ 칠러 A 정지 → 예비 C 투입 / 예비 사용 불가 시 용량·온도 경고
+ 로컬 이미지·SVG 글꼴·CSS·JS 내장 / 단일 HTML 오프라인 동작
```

기존 구역 번호, 8비트 탑다운 지도, 방향키/WASD·Enter/Space/Z·D-pad+A, AUTO/모바일/컴퓨터 모드, 탐험 노트 저장, 야간 모드, 세 랙 참고 화면, 카세트·Interior Explorer와 BOM/부품/Supply Chain 링크를 유지합니다. 시설 배치 면책 문구도 유지합니다. 기존 Gray/White 참고 비트맵은 재분류와 모순되는 MMR/Carrier 그림을 포함하므로 구역 가이드에서 숨깁니다. 랙별 그림은 유지하며 모두 내장합니다.

## 예시 가정

모든 교육 패널에 **예시 가정값(20MW급 AI 캠퍼스 기준)**을 표시합니다. 여기서 20MW는 IT 설계 부하 예시입니다.

- 정상상태 PUE 1.40 가정에서 시설 전력 28MW. 수전 설계 규모 40MVA는 별도의 예시이며 전압·역률·허용전류·예비율을 산정한 확정 설계가 아닙니다.
- 공랭 랙 10~20kW, AI 랙 60/100/120kW 예시 등급.
- N+1 UPS: 5MW × (4 운전 + 1 예비). 2N 비교: A/B 각각 20MW. N+1과 독립 A/B 경로는 다른 개념입니다.
- 정전 시나리오: UPS 배터리 300초, 발전기 정격 확보 30초, ATS 전환 40초, UPS 정상 입력 확인 45초. 실제 1초를 시나리오 5초로 압축합니다.
- 냉각: CW 29/34°C, FWS 18/24°C, TCS 30/40°C. 냉각탑 계통과 시설 냉수는 칠러 냉매 계통을 통해 열을 주고받고, FWS/TCS는 CDU 열교환기로 분리합니다. 같은 물이 냉각탑부터 GPU까지 흐르는 도식이 아닙니다.
- 냉각 장애: IT 열부하 20MW에 칠러 10MW × 2 운전 + 1 예비를 두는 교육 시퀀스. 부대 열부하·냉매/수리·배관 설계 계산은 포함하지 않습니다. 예비 투입 30초, 예시 TCS 공급 30→32→30°C; 예비 사용 불가 시 10MW 용량 부족과 경고가 지속됩니다. 이 온도/시간은 물리 동특성 계산이 아닌 교육용 스크립트입니다.
- 광심 연습은 링크 수×심수×추가 예비율로 계산하며 MW에서 환산하지 않습니다. 20% 예비율은 편집 가능한 예시입니다.

## 검증

```sh
NODE_PATH=/path/to/node_modules node DCI/DataCenter/tests/campus_onboarding_browser.cjs
NODE_PATH=/path/to/node_modules node DCI/DataCenter/tests/campus_connected_racks_browser.cjs
```

Playwright Chromium 141. 신규 테스트는 `file://`로 단일 HTML을 열고 로컬 파일/외부 라이브러리 요청 없이 동작하는지 확인합니다. 58개 발견 항목 접근 가능성, 모든 상세 패널, 7개 역할 패널 및 Switchgear/ATS의 Gray Zone 단일 배치, 조사·노트·재시작 저장, 문제 풀이·광심 입력 변경, 실제 키보드로 구역 간 전력 경로 걷기, D-pad/기존 키, 모달 정지, 야간/화면 모드, 액냉/광 탭 전환, 1440/768/390/320px 가로 넘침과 JS 오류를 검증합니다. 가상 애니메이션 시계로 정전·N+1 복구·예비 사용 불가·정상 복귀를 확인합니다. 기존 랙 테스트는 원래 세 그림·확대·키보드 조사·노트 재열기를 검증합니다.

## 근거

- DOE: https://www.energy.gov/cmei/femp/cooling-water-efficiency-opportunities-federal-data-centers
- NVIDIA: https://networking-docs.nvidia.com/cablemanagfaq/cable-and-connector-definitions
- Vertiv UPS: https://www.vertiv.com/en-cn/about/news-and-events/articles/product-class-articles/network-server-and-storage-uninterruptible-power-supply-ups/

4번 ‘광 매체, 네트워크’의 세부 요구는 아직 전달되지 않았으므로 기존 광심 수 학습과 이번 구역 재분류를 유지합니다. 전체 4번 신규 기능 완료를 의미하지 않습니다. 학습 기록은 브라우저 localStorage이며 LMS 수료나 현장 운전 자격이 아닙니다.
