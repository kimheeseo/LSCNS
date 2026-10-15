# STEP 15 / 20 — 독립 MMR 광회선 인입/교차접속 공간 정밀화 보고서

- 기준: 2026-10-15 KST · `kimheeseo/LSCNS` / `main`
- 범위: **STEP 15만 수정**. 2D Gold Pixel Tour/BOM 소스 및 STEP 16 Co-location 모듈은 변경하지 않음.
- 상태: **3D 코드 구현 + 모의 검증 완료 / 실 Chromium·WebGL·모바일 시각 검증 미실행**
- 테스트 모드: `LS_Datacenter_3D.html?engine=three`

## 1. 기존 구현과 개선 근거

`mmr-room.mjs`와 `zone-specializations.mjs`에 이미 4개 구획·메뉴·광회선 개념 경로가 있었으나 다음 문제가 있었다.

1. 기존 Carrier/ODF/DWDM/Core의 앞면이 서로 유사한 추상 패널로 구성되며 일부 모듈은 실제 호스트 장비 앞면에서 떨어져 배치됨.
2. 구획별 실제 3D 출입문 개폐 객체가 없었음.
3. 광 링크가 선택돼도 개별 3D Tube가 raycast 대상이 아니었음.
4. 회선 필터는 비선택 경로를 완전히 숨겼고, 다시 빌드했을 때 강조 선택이 초기화될 수 있었음.
5. 출발·도착 장비와 케이블 개념 설명이 선택 패널에 함께 나타나지 않았음.

## 2. 구현 내용

| 대상 | 구현 |
|---|---|
| **Carrier Entrance / ER** | 별도 입구 구획, Carrier Demarc 판넬과 케이블 인입 gland 4개, 전면 24개 포트 예시 |
| **MMR / DWDM Handoff** | 기존 MMR 자산 랙의 앞면에 DWDM/OTN-style muxponder 4개 모듈과 포트 32개를 별도 3D 구성 |
| **ODF / Cross-connect** | 전면 patch-board와 3단 Optical Patch Shelf, 각 8개 어댑터 표현 |
| **Core / WAN Router** | 개별 interface card 3개, 각 8개 포트 cage와 서로 다른 상태색 |
| **룸 분리** | 기존 Network/MMR 구역 안 4개 물리 구획·투시형 구분벽 보존, 각각 독립 `THREE.Group` 기반 힌지 도어·카드리더 추가 |
| **접속 경로** | 외부 Carrier → Carrier Demarc → ODF → DWDM Handoff → Core를 4개 구간으로 표현; 4개 랙에서 광 Raceway로 내려오는 3D 서비스 drop 추가 |
| **광 링크 선택** | 각 링크의 `THREE.TubeGeometry`를 클릭 가능한 raycast 대상과 메타데이터(`mmrSegment`)로 등록 |
| **강조 동작** | 선택한 광 경로 opacity 0.95, 관련 없는 광 경로 opacity 0.10. 나머지 경로를 삭제하지 않아 공간 맥락 보존 |
| **조작** | MMR 메뉴에 4개 출입문 버튼, 링크별 출발 장비·도착 장비·링크 종류 설명 추가 |
| **상태 보존** | 구역/단면 변경으로 Scene이 재생성될 때 회선 강조 필터 유지, 오래된 동적 곡선 geometry·material 해제 |

도어는 기존 Three.js 렌더 사이클에서 `tick(dt)`로 실제 `THREE.Group.rotation.y` 값을 변경한다. 단순한 이미지 교체나 2D 플레인이 아니다. 실제 장비 모형이 아닌 **교육용 형상/좌표/연결 예시**이며, 특정 Carrier POP·DWDM 제조사·제품 포트 수·광 링크 속도/손실을 입증하지 않는다.

## 3. 검증 범위와 결과

### 실행한 소스 기반 모의 검증

- **MMR 형상·링크/도어 테스트 16/16 PASS:** 4 구획, 4 문, 총 128개 패치·DWDM 전면 포트 형상(교육용), 4개 내부 서비스 드롭, 4개 지정 광회선, 독립 장비 모델 종류, 선택 metadata, 3D 클릭 메타정보, 문 그룹 회전, 필터 전환, 재생성 상태 보존, 자원 해제, 회사 server-room 분리.
- **Three Scene Bridge 연동 8/8 PASS:** `zone-specializations.mjs` 실제 GitHub 소스 로드, 랙 선택 연결, 클릭 경로 ID 식별, 카메라 구역 전환 시 노출 제어, 재진입 시 선택 상태 유지, 서버실 모드 격리.
- **커밋한 Node 호환 테스트 소스의 실제 JS 모의 실행 4/4 PASS:** `step15.unit.test.mjs`를 GitHub에서 다시 읽고 구현 모듈과 함께 자바스크립트 실행기로 실행함. `node --test` CLI 환경에서의 실행 로그는 별도로 생성하지 않음.
- Scene/테스트 파일 구문 파싱은 정상.

### 작성했으나 실행하지 않은 검사

`step15.playwright.spec.mjs`에 실제 Chromium에서 검증하도록 다음 테스트를 추가했다: Carrier·DWDM·ODF·Core 형상, 물리 도어의 회전, 광회선 선택 시 투명도 및 경로 정보, server-room 격리, 390px 모바일 화면/스크린샷, 브라우저 에러.

**현재 결과는 GPU/WebGL 실제 렌더링 정상 판정이 아니다.** 공개 GitHub Pages에 대한 독립 Chromium 접속이 이 세션에서 불가능하여 실제 새 화면 캡처·GPU 메모리·FPS·플랫폼 별 픽킹은 미확인이다.

### 개발 환경 재시험

```bash
node --test DCI/DataCenter/ls3d-three/step15.unit.test.mjs
npm install --no-save @playwright/test
npx playwright install chromium
LS3D_BASE_URL=https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html npx playwright test DCI/DataCenter/ls3d-three/step15.playwright.spec.mjs
```

## 4. 변경/신규 파일

- `ls3d-three/mmr-room.mjs` — MMR 3D 랙 전면/히든 공간/문/광회선 모형
- `ls3d-three/zone-specializations.mjs` — 독립 문 조작/회선 클릭/요약 UI, 모듈 버전
- `ls3d-three/scene-adapter.mjs` — STEP 15 버전 모듈 import 경로 갱신
- `ls3d-three/bootstrap.mjs` 및 `LS_Datacenter_3D.html` — GitHub Pages ES module 캐시 버전 갱신
- `ls3d-three/step15.unit.test.mjs` — 4개 소스 기반 Node 호환 회귀 테스트
- `ls3d-three/step15.playwright.spec.mjs` — 4개 브라우저 수용성 테스트(미실행)
- `architecture/STEP15_MMR_REPORT.md` — 본 보고서
- `LS3D_V4_20_TASK_ROADMAP.md`, `LS3D_PROGRESS_STATUS.md` — STEP 15 상태 반영

기존 보호 파일 SHA(변경 없음):
- `DCI/DataCenter/LS_Datacenter_Campus.html`: `9c84f24f2101043d7c113ce2d0da50e7245bc318`
- `DCI/DataCenter/index.html`: `692767f0c234df3b8310642e3acd032137c68baa`

## 5. 제약/남은 검증

1. MMR 외벽/분리 구획은 실제 건축 시방이 아니라 교육용 레이아웃이며 출입통제 설계 승인이나 피난 동선 검토가 아님.
2. 포트/광회선은 시연용 예시로서 용량·광손실·실제 DWDM 구성·실 케이블 부하 확인을 포함하지 않음.
3. 실제 클릭 좌표·카메라 가림·문과 랙 간섭·모바일 화면·브라우저 콘솔은 19단계 통합 QA 및 수동 확인 대상.
4. 15단계 완료는 **구현/모의 검증 기준**이며 실제 브라우저 수용성 완료 판정은 보류.

- 3D 실행: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three
