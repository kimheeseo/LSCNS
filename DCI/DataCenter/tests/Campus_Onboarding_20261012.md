# Campus Onboarding & OJT · 2026-10-12

Campus Tour에 60분 입문 과정을 추가했습니다. 기존 픽셀 탐험과 랙 상세 그림을 유지하며 교육 과정과 탐험 진행을 별도로 표시합니다.

| 단계 | 목표 | 권장 시간 |
|---|---|---:|
| 1 큰 그림 | 전력·열·데이터 흐름과 단위 구분 | 5분 |
| 2 전력 | UPS·발전기·배전·A/B 잔여 용량 | 10분 |
| 3 냉각 | 공기·IT수·시설수 루프, 열교환 | 10분 |
| 4 랙 | Trunk·Cassette·Optics·ToR·전원 연결 | 10분 |
| 5 광 | 전력과 광심 수 구분, 경로별 심수 연습 | 15분 |
| 6 운영 | 상태·영향·연관 계통 기록과 보고 | 10분 |

- 학습 목표, 용어 사전, 장비 그림 바로 열기, 구역 이동, 단계별 객관식 확인과 설명을 제공합니다.
- 강의 모드는 본문과 그림을 크게 표시하고 배경음을 끕니다. 자유 탐험 모드에서는 기존 7개 구역을 탐험할 수 있습니다.
- 정답 확인한 단계만 완료로 기록합니다. 단계 선택은 복습과 강의를 위해 자유롭게 할 수 있습니다.
- `ls-campus-onboarding-v1`과 `ls-pixel-discoveries`는 브라우저 localStorage 기록입니다. 초기화 확인창에서 두 기록을 지웁니다. 저장 실패 시 이번 방문에서만 기록함을 표시합니다. LMS나 수료증은 제공하지 않습니다.
- 배경음은 기본 꺼짐이며 버튼으로 켜고 끕니다.
- 심수 계산: `ceil(활성 링크 수 × 링크당 사용심 × (1 + 추가 예비율))`. 예비율 20%는 교육 예시입니다. A/B 독립 경로는 각각 산정하고 커넥터 미사용 위치·극성·손실·거리·포트 규격·케이블 모듈 단위는 별도 확인합니다. 심수 충족은 전체 호환성 PASS가 아닙니다. W/F 고정 환산은 하지 않습니다.
- 물 루프를 열 이동 화살표와 구분했습니다. 보충수 화살표는 저장조에서 냉각탑으로 향합니다. 구역 번호를 메뉴와 통일했고 캠퍼스 북측 통로로 CHILLER BANK에 접근할 수 있습니다.
- 모든 dialog가 지도 이동을 일시 정지합니다. 학습 입력과 버튼 조작에 이동 단축키가 간섭하지 않습니다.
- 온보딩에서는 설비와 무관한 일반 랙 참고 그림을 숨기고 해당 설비 개념도를 엽니다. 자유 탐험의 참고 이미지는 '참고 일러스트'로 표시합니다.
- 한글 글꼴은 Noto Sans KR의 필요한 글자만 포함한 WOFF입니다. 라이선스: `assets/campus-tour/connected-racks-font-OFL.txt`.

## 검증

```sh
NODE_PATH=/path/to/node_modules node DCI/DataCenter/tests/campus_onboarding_browser.cjs
NODE_PATH=/path/to/node_modules node DCI/DataCenter/tests/campus_connected_racks_browser.cjs
```

Playwright Chromium 141에서 1440/768/390/320px 화면을 확인했습니다. 온보딩 테스트는 정오답 처리, 6단계 완료, 새로고침 저장, 초기화 취소/확정, 광심 계산 입력 변화·유효성, 전체 강의 장비 그림, 42개 발견 항목 접근 가능성, 모달 중 이동 정지, 모드 전환, 한글 글꼴, 가로 넘침 및 JS 오류를 확인합니다. 기존 테스트는 세 랙 그림의 차이, 실제 키보드 조사, 노트 재열기, 확대와 모바일 동작을 확인합니다.

## 강의 근거

- [DOE · Cooling Water Efficiency](https://www.energy.gov/cmei/femp/cooling-water-efficiency-opportunities-federal-data-centers)
- [NVIDIA · H100 Network Fabrics](https://docs.nvidia.com/dgx-superpod/reference-architecture-scalable-infrastructure-h100/latest/network-fabrics.html)
- [NVIDIA · Cable and Connector Definitions](https://networking-docs.nvidia.com/cablemanagfaq/cable-and-connector-definitions)

배치와 도식은 교육용이며 실제 LS 사업장 배치·운전 절차·제품 사양을 나타내지 않습니다. OJT 마무리는 강사와 사업장 문서를 함께 확인합니다.
