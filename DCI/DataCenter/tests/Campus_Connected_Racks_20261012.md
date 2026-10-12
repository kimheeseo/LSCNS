# Campus Gold Tour · 연결 랙 디자인 (2026-10-12)

사용자 제공 참고 이미지 4의 청록 광 점퍼 연결 느낌을 데이터 홀의 저해상도 Canvas 랙에 반영했다.
포트 쌍, 계단형 여장 루프, 측면 케이블 번들, 상부 인입을 표시하며,
COMPUTE B는 고밀도 정면 모듈로 구분한다. 캠퍼스/룸 구조와 이동 경로는 그대로 사용한다.

| 조사 위치 | 참고 이미지 | 새 조사 그림 |
|---|---|---|
| ① GPU ROW A | 3번 다중 랙 도면 | 8캐비닛 Supercluster, 상부 트레이, 랙별 인입, Fabric/NIC 연결, A/B 전력 경로 |
| ② GPU ROW B | 5번 랙 후면 사진 | 보라·청록 점퍼 루프, 식별 라벨, 반복 포트 뱅크, 주황 측면 번들·수직 관리 |
| ③ COMPUTE B | 6번 장비 정면 사진 | 2랙 고밀도 모듈 뱅크, 반복 슬롯, 주황 래치·파란 상태 표시, 측면 케이블 인입 |

각 위치에 별도 조사 ID를 부여해 같은 일반 랙 그림이 반복되지 않도록 했다.
랙 앞에서 A/Z/Enter로 조사하거나 DATA HALL 아래의 ①②③ 바로 열기 버튼을 사용한다.
`전체 보기`는 화면 폭에 맞추며 `세부 확대`는 그림 내부 스크롤을 지원한다.
탐험 노트에서도 각 그림을 다시 열 수 있다.

SVG는 직접 작성한 픽셀 스타일 개념 그림이다. 참고 사진의 제품 모델·정확한 수량·포트 규격을 추정하지 않는다.
기존 ToR 카세트, NETWORK A/B, MMR, ODF 조사 화면은 유지한다.
새 한글 설명은 Noto Sans KR의 추가 subset WOFF로 제공한다.
원본 글꼴: https://github.com/google/fonts/tree/main/ofl/notosanskr
라이선스: `assets/campus-tour/connected-racks-font-OFL.txt`.

## 검증

```
NODE_PATH=/path/to/playwright/node_modules node DCI/DataCenter/tests/campus_connected_racks_browser.cjs
```

- Chromium 141, PC 1440×1100 / 모바일 390×844
- 세 위치의 서로 다른 제목·SVG 로드와 카세트 버튼 숨김
- 확대/전체 보기 상태와 내부 스크롤 폭
- 실제 방향키로 각 랙까지 이동한 뒤 A 버튼으로 조사
- 탐험 노트의 발견 항목에서 동일 그림 재열기
- 모바일 이미지 로드·확대·닫기, 페이지 가로 넘침 없음
- 다른 구역에서 랙 바로 열기 메뉴 숨김
- 페이지 JavaScript 오류 0건
- 화면 캡처에서 글꼴과 케이블·모듈 표현 검토
