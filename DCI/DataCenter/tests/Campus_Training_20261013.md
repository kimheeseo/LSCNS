# Campus 주니어·시니어 모드 보완 — 2026-10-13 KST

대상: `LS_Datacenter_Campus.html`. 기준 main: `4d0e7b4f75702c995fc58c8e4c320b138986711c`.

## 변경 전 구조

- 단일 HTML: Canvas 384×320, 타일 16px, `maps`의 grid/objects/doors와 DOM 설명창.
- 주니어: 7구역·12개 대표 설비, 쉬운 설명. 시니어: 73개 상세 탐험 항목.
- `ls-campus-training-mode`에 모드, `ls-campus-junior-discoveries-v1`에 주니어 발견 기록. 시니어 발견 기록은 기존 `ls-pixel-discoveries`에 별도 저장.
- 모드 선택 바는 이미 console 내부로 이동되어 main의 두 열 배치를 유지함.

## Diff: 이번 변경만

```diff
- 주니어 설명창의 높이 제한이 없어 390px 화면에서 지도 밖으로 약 344px 넘침
+ 지도 안의 96% 높이에서 내부 스크롤, 상단 닫기 버튼 유지
- 새 주니어 설명과 제목이 OS 한글 글꼴에 의존
+ 실제 문구의 한글 495자 모두 포함한 Noto Sans KR WOFF를 HTML에 내장
- 주니어 안내와 구역 버튼이 모바일에서 좁은 두 칸에 섞임
+ 안내는 화면 너비 사용, 구역 목록은 별도 두 열
+ 아직 조사하지 않은 다음 설비의 위치로 이동하는 버튼과 상태 안내
+ A 조사창에도 다음 설비 위치 이동 버튼 추가
+ 출입문에 목적지 표시, 문을 바라보면 이동할 구역 안내
- 시니어 회귀 테스트가 과거 58항목 및 이전 기본 모드를 기대
+ 시니어 명시 선택과 현재 73항목을 기대하도록 갱신
+ 주니어/시니어 분리·모바일 창·다음 안내·문 이동 회귀 테스트 추가
```

위치 이동은 발견 처리하지 않음. 기존 A/Enter/Space/Z 조사로만 기록. 12개 발견 후 다음 버튼은 완료 상태로 전환하며 탐험 노트에서 다시 읽을 수 있음. 설명창에 키보드 포커스가 있으면 방향키는 내용 스크롤에 사용되며 게임은 정지 상태를 유지.

그림은 기존 픽셀 팔레트와 Canvas 코드만 사용. 생성 이미지·외부 라이브러리·필수 외부 자산 요청 없음. 내장 폰트는 기존 SIL OFL 고지 유지. 기존 야간 모드·랙 Interior Explorer·BOM/부품 DB/Supply Chain 링크·교육용 면책 문구 유지. 새 수치 설비 모델은 추가하지 않음.

## 검증

Playwright 1.56.1 / Chromium 141 (build 1194), 로컬 file://에서 검증.

```sh
NODE_PATH=/path/to/playwright/node_modules node DCI/DataCenter/tests/campus_training_browser.cjs
NODE_PATH=/path/to/playwright/node_modules node DCI/DataCenter/tests/campus_onboarding_browser.cjs
```

- 주니어 12개를 다음 안내→A 조사로 발견, 안내만으로 발견되지 않음.
- 주니어 12/12·시니어 별도 기록, 모드 전환·새로고침·노트 재열기·야간 모드 유지.
- 7구역의 문 표시가 실제 doors 목적지와 일치, 방향키로 Gray→White 진입.
- 1440/768/390/320px 페이지 너비, 설명창 지도 내 배치, 작은 화면의 내부 스크롤과 닫기.
- 시니어 73개 항목 도달 가능성, 6개 OJT 콘솔 조사/퀴즈 저장, 광심 입력 변경 계산, 3개 랙 액냉/광 보기.
- 실제 방향키로 캠퍼스→Utility→Gray→White 전력 경로와 PDU B 분기 통과.
- UPS 즉시 보완·발전기/ATS 순서, 냉각 N+1 복구·예비 칠러 사용 불가 경고 시나리오.
- 외부 자산 요청 0, JavaScript 오류 0. 한글 글리프 누락 0.

배포 후에도 동일 테스트를 실행할 수 있음:

```sh
CAMPUS_URL=https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_Campus.html NODE_PATH=/path/to/playwright/node_modules node DCI/DataCenter/tests/campus_training_browser.cjs
```
