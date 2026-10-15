# STEP 16 / 20 — Co-location 고객 A/B/C 물리 격리·케이블 3D 모델 정밀화

- 수행 기준: 2026-10-15 KST, `kimheeseo/LSCNS`, `main`
- 구현 판정: **코드·모의 검증 완료 / 실제 Chromium·WebGL·모바일 실화면 수용 검증 대기**
- 사용자 승인 범위: **STEP 16만 진행**. STEP 17~20 구현은 하지 않음.
- 실행 페이지: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three

## 1. 기존 구현과 실제 문제

STEP 16의 원래 `colocation-cages.mjs`는 3고객 구획, Mesh Fence, 도어 그룹, Fiber 및 PDU A/B 광·전원 연결을 가지고 있었다. 그러나 다음 개선점이 남아 있었다.

1. 공간 너비가 좁으면 원래 고객 랙 위치를 무시하고 **동일한 3등분 펜스**로 되돌아가 랙을 관통할 수 있었다.
2. 펜스가 반투명한 큰 평판과 몇 개의 막대로 구성되어 실제 선재 메시의 3D 밀도가 부족했다.
3. 광/전력 경로는 `visible=false`로 완전히 숨겨졌고, 3D 케이블 자체의 raycast 식별 및 회선 설명이 어려웠다.
4. `showSection` 전환으로 씬을 재생성하면 출입문의 개폐도가 초기화됐다.
5. 고객별 실제 배선 인터페이스·전력 A/B 피더·광 어댑터가 시각적으로 분명하지 않았다.

## 2. 이번 구현 결과

| 구분 | 구현 결과 |
|---|---|
| 고객 구획 배치 | 원본 자산 `colo` 및 `colo-rack-a/b/c`의 실제 좌표와 폭에서 경계 계산. 랙 중첩·Cage 외곽 초과 또는 폭 부족 시 임의 3등분을 하지 않고 `layoutIssue` 보고 / `ready=false` |
| 개방형 Mesh Fence | 4개 세로 분리 경계를 가로/세로의 실제 얇은 3D steel members로 구성. 고객 구획별 후면 및 전면 양 측면에도 반복 3D 와이어 추가 |
| 고객 A/B/C | 분리된 `THREE.Group`·고유 출입문/이름표·랙 관리. 기존 BOM 기준 고객 랙 ID 보존 |
| 출입 | 3개 힌지 애니메이션 `rotation.y` 작동. 구역 모델 재생성 시 기존 열림 목표값과 각도 보존 |
| 광 | 고객당 독립 3D Fiber Raceway, 8개 시각용 어댑터, 1개 `THREE.TubeGeometry` 패치 경로 및 메타데이터 |
| 전력 | 고객당 Rack PDU A/B 피더 형상과 A/B 별 2개 독립 3D 케이블 경로 |
| 클릭·강조 | 고객별 광·PDU A·PDU B 개별 선택 및 전용 표시. 선택 경로 opacity 0.95, 나머지 opacity 0.11 — 객체는 계속 존재 |
| 연결 정보 | 고객, 구획, 출발 공용 경계(개념), 도착 랙, 경로 종류 노출. `verified:false`, 미확인 커넥터/용량 명시 |
| 원본 보호 | 2D Tour/BOM, 15단계 MMR 원본 `mmr-room.mjs` 미수정, 기존 7구역·랙/카메라 인터페이스 유지 |

**표시하는 어댑터 수/피더/경로/높이는 개념 모델로서 실제 발주 BOM·제조사 패널 인터페이스·고객 전력 이중화 성능·광손실 확인이 아니다.**
MMR↔고객 Cage 사이의 캠퍼스 전체 End-to-End 경로는 STEP 17 작업 범위로 남는다.

## 3. 수행 검증

### 실제 소스 기반 Three 호환 모의 검사

- **28/28 PASS:** `colocation-cages.mjs`, `zone-specializations.mjs`, `mmr-room.mjs` 실제 GitHub 코드를 로드 후 실행.
- 포함: 고객 A/B/C 좌표·랙 충돌없는 경계, 150+ 메시 선재, 3개 Gate, Fiber 3개 + Power 6개 독립 경로, 독립 광 어댑터/후면 전력 A/B 인터페이스, 9개 3D TubeGeometry의 pick metadata, 회선 선택 opacity, 문 각도 회전, rebuild 상태 보존, 랙 충돌 시 모델 실패 표기, geometry.dispose, server-room 및 구역 변경/복귀 시 상태 보존, Scene Bridge 개별 회선 클릭.
- **재실행 가능한 커밋 테스트 `step16.unit.test.mjs`: JS 자체 실행 5/5 PASS.** 테스트는 THREE API 호환 mock을 사용하며 GPU draw call, 픽셀, 충돌 물리엔진, 제조사 성능 검증은 아니다.
- 단일 라우트 키 매칭 오류를 구현 중 발견해 `power-A`/`power-B` 키로 정규화 후 전체 검사 통과.

### 작성했지만 실행하지 않은 브라우저 테스트

`step16.playwright.spec.mjs` 4개 테스트로 다음 항목을 확인하도록 구현했음:
- 고객 A/B/C 실제 3D Mesh 및 랙 위치, 개별 회선과 파워 이중화
- 특정 고객 PDU B 선택 강조 후 나머지 경로 반투명화
- 선택한 고객 출입문만 3D 힌지 회전
- 모바일 390×844 화면 캡처, server-room 격리, `?engine=legacy` 회귀

**실제 Chromium/WebGL 테스트 및 스크린샷·FPS는 이번 세션에서 실행하지 않음. 브라우저 PASS로 주장하지 않는다.**

검증 실행:
```bash
node --test DCI/DataCenter/ls3d-three/step16.unit.test.mjs
npm install --no-save @playwright/test
npx playwright install chromium
LS3D_BASE_URL=https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html npx playwright test DCI/DataCenter/ls3d-three/step16.playwright.spec.mjs
```

## 4. 변경 파일 / 배포

- `ls3d-three/colocation-cages.mjs` — 고객 물리 펜스, 랙 footprint 검증, 배선 메타데이터, 클릭 가능 Tubes, 상태 보존
- `ls3d-three/zone-specializations.mjs` — 고객별 9개 상세 회선 메뉴·클릭 처리·요약 정보
- `ls3d-three/scene-adapter.mjs` — 새 Bridge 버전 참조
- `ls3d-three/bootstrap.mjs`, `LS_Datacenter_3D.html` — 최신 모듈을 읽도록 버전 경로 갱신
- `ls3d-three/step16.unit.test.mjs` — 소스 기반 테스트
- `ls3d-three/step16.playwright.spec.mjs` — Chromium 수용 테스트 초안
- `architecture/STEP16_COLOCATION_REPORT.md` — 본 문서
- `LS3D_PROGRESS_STATUS.md` 및 `LS3D_V4_20_TASK_ROADMAP.md` — 진행 기록

보존된 기존 파일 SHA:
- `LS_Datacenter_Campus.html`(2D): `9c84f24f2101043d7c113ce2d0da50e7245bc318`
- `index.html`(BOM): `692767f0c234df3b8310642e3acd032137c68baa`

## 5. 추후 필요한 실제 검증

1. Chrome/Edge에서 Data Hall 선택 → 고객 A/B/C 위치가 펜스 내부에 놓이는지 3D 화면 확인.
2. 광/전력 A/B 필터, 출입문 조작, 원래 서버 인출·랙 클릭과 기능 충돌 없는지 확인.
3. 모바일 터치, 계측 FPS/메모리, 카메라 벽 관통/가림 여부 확인.
4. 17단계에서 MMR→고객 Cage, 공용 전원·냉각 인입 등 공통 배관/회선 경로를 실제 지오메트리로 연결하되 지금의 교육용 회선과 구별.
