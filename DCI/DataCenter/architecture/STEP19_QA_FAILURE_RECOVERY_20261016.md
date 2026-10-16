# LS Datacenter 3D — Step 19 GitHub Actions 반복 실패 원인 및 해결 검증 보고서

**작성/검증 정리일:** 2026-10-16 (KST)  
**대상 저장소:** [kimheeseo/LSCNS](https://github.com/kimheeseo/LSCNS) · `main`  
**대상 웹 페이지:** [LS Datacenter 3D (Three.js)](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three)  
**범위:** Step 19의 두 GitHub Actions 통합 품질검사 및 관련 GitHub Pages 배포. 실측 시설 설계 적합성 평가 보고서는 아님.

## 1. 최종 검증 결과 (GitHub Actions에서 직접 확인)

| 구분 | 결과 | 실행 번호 / 커밋 | 실제 로그 |
|---|---|---|---|
| **Step 19 Integrated QA** | **SUCCESS · 브라우저 테스트 9/9 PASS** | #57 · `7e3b5b9f37de06598da3a440198440ced3a5a7b6` | [실행 #57](https://github.com/kimheeseo/LSCNS/actions/runs/38820787435) |
| **Step 19 Three.js Integration QA** | **SUCCESS · 브라우저 테스트 5/5 PASS** | #11 · `172d3fc6b16637e659cf5ec423a2c5a4332a4e9e` | [실행 #11](https://github.com/kimheeseo/LSCNS/actions/runs/38819653752) |
| **GitHub Pages** | **SUCCESS · 배포 작업 성공** | #937 · `7e3b5b9f37de06598da3a440198440ced3a5a7b6` | [배포 #937](https://github.com/kimheeseo/LSCNS/actions/runs/38820786560) |

주의: 위 5/5 및 9/9는 **서로 다른 시점의 커밋**에서 각각 성공한 두 별도 QA 스위트입니다. 두 숫자를 동일한 커밋의 단일 14/14 재검증으로 해석하지 마세요. 또한 GitHub Pages **배포 워크플로의 성공**은 모든 실제 휴대폰/브라우저에서의 동작 및 FPS 성능 보증이 아닙니다.

## 2. 오류 이메일 반복 원인

- GitHub Actions에서 `push`마다 관련 경로 변경 시 QA를 재실행하도록 설정되어 있었음. 한 번의 실패 이메일 이후 후속 수정 커밋에서도 테스트가 실패하여 새로운 실패 알림이 계속 발생.
- **첫 방문 가이드 오버레이**: `#guide.twin-modal.open`이 실제 DOM 클릭/터치 이벤트를 가로채 랙 후면 카메라, MMR Carrier 출입문, Co-location C 출입문, 랙 패널 펼치기 등이 `intercepts pointer events`로 시간 초과.
- **짧은 화면의 랙 패널**: 부유 랙 조작 패널이 화면 경계를 넘어서 상단 KPI/지표 영역과 겹치고 `locator.click`이 차단됨.
- **모바일 제스처 테스트**: 화면 상단의 고정된 좌표에 두 손가락 터치를 전송하면, 모바일 화면의 지표/안내 영역이나 뷰포트 밖에 닿아 3D 캔버스에 이벤트가 도달하지 않았음.
- **서버실 전환 판정 오류**: 3D 구획 객체의 `ready`는 '초기화 완료'를 뜻하므로 서버실에서 숨김 처리됐더라도 true일 수 있음. 실제 표시 상태 `root.visible`을 검사해야 함.
- **두 번째 QA 서버 경로**: 작업 디렉터리가 저장소 루트가 아니어서 로컬 페이지가 HTTP 404를 반환.
- **QA 결과물 경로**: Playwright 출력 경로를 중복 접합하여 아티팩트 업로드가 누락되던 문제.

## 3. 수정 방법과 적용 코드

| 대상 파일 | 적용 조치 | 기대 효과 |
|---|---|---|
| `ls3d-three/step19.integration.playwright.spec.mjs` | 테스트에서 첫 방문 안내창을 닫고, 실제 화면에 닿는 모바일 터치 좌표 검색 및 사전 스크롤; 서버실 실제 표시 상태 확인 | 랙/MMR/Cage 조작·모바일 pan/pinch·시설 격리 테스트 |
| `ls3d-three/step19.integration.spec.mjs` | 다른 QA에서도 안내창 닫기; `ready` 대신 실제 `root.visible` 검증 | 별도 Chromium QA 오판 방지 |
| `ls3d-three/playwright.step19.config.mjs` | 서버 `cwd`를 저장소 루트로 변경하고 결과물 경로 수정 | HTTP 404 및 아티팩트 경로 수정 |
| `ls3d-three/scene-viewport-policy.mjs` | 작은 데스크톱 화면에서 랙 패널 높이를 실제 뷰포트 내로 제한 | 랙 조작 버튼이 다른 UI 밑으로 넘어가지 않도록 배치 |

관련 수정 커밋(시간순):

1. [`fe7848a`](https://github.com/kimheeseo/LSCNS/commit/fe7848a46f2d6d8a2b3b540480f782258a6bdf41) — 통합 QA 첫 방문 화면 초기 처리
2. [`7f39803`](https://github.com/kimheeseo/LSCNS/commit/7f39803120e7e05f6378a2474b1ce079a849d358) — 다른 QA의 localhost 루트 수정
3. [`22d4110`](https://github.com/kimheeseo/LSCNS/commit/22d41105e4bb93d9510948ac29d547b0f2c9be2b) — 다른 QA 온보딩 처리
4. [`c10f54e`](https://github.com/kimheeseo/LSCNS/commit/c10f54e4d45b8817e6bdf603a70946d0131bfb75) — 서버실 상태 판정 보완
5. [`172d3fc`](https://github.com/kimheeseo/LSCNS/commit/172d3fc6b16637e659cf5ec423a2c5a4332a4e9e) — 화면 안내창 실제 닫기 및 3D 객체 표시 여부 검사
6. [`ae7c86f`](https://github.com/kimheeseo/LSCNS/commit/ae7c86f1cce57a3cbd0ad7f7d051ed0daad8fd1d) — 랙 패널 높이 및 모바일 캔버스 대상 좌표
7. [`7e3b5b9`](https://github.com/kimheeseo/LSCNS/commit/7e3b5b9f37de06598da3a440198440ced3a5a7b6) — 모바일 3D 캔버스를 스크롤한 뒤 제스처 테스트; **9/9 통과**

이 수정에서 `LS_Datacenter_Campus.html`(기존 2D Gold Pixel Tour) 및 DataCenter BOM 계산 엔진을 의도적으로 변경하지 않았습니다. 사용자의 온보딩 안내 기능은 제품 화면에서 유지합니다.

## 4. 단계별 회귀 추적

| 실행 | 결과 | 주요 확인 사실 |
|---|---|---|
| [#49](https://github.com/kimheeseo/LSCNS/actions/runs/38774347666) | FAIL, 9개 중 5개 실패 | `#guide` 간섭 및 모바일 실패 시작점 |
| [#54](https://github.com/kimheeseo/LSCNS/actions/runs/38818116292) | FAIL, 9개 중 6개 실패 | 단순 localStorage 사전 설정만으로는 가이드창이 닫히지 않음 |
| [#10 — 별도 QA](https://github.com/kimheeseo/LSCNS/actions/runs/38818116255) | FAIL, 5개 중 2개 실패 | 가이드 간섭·서버실 상태 검사 |
| [#11 — 별도 QA](https://github.com/kimheeseo/LSCNS/actions/runs/38819653752) | **SUCCESS, 5/5** | 첫 번째 통합 수정 배치 효과 확인 |
| [#55](https://github.com/kimheeseo/LSCNS/actions/runs/38819653745) | FAIL, 9개 중 2개 실패 | 랙 패널/지표 겹침 및 모바일 두 손가락 동작 |
| [#56](https://github.com/kimheeseo/LSCNS/actions/runs/38820390909) | FAIL, 9개 중 1개 실패 | 랙 문제 해결; 모바일 캔버스 영역 확보 실패 |
| [#57](https://github.com/kimheeseo/LSCNS/actions/runs/38820787435) | **SUCCESS, 9/9** | 최종 모바일 스크롤 처리 후 전부 통과 |

Actions HTML 리포트/스크린샷/trace 등은 각 실행의 **Artifacts**에서 확인할 수 있으며 보존 기간 경과 후 자동 삭제될 수 있습니다. 기록은 GitHub 실행 링크 및 이 문서를 기준으로 유지합니다.

## 5. 검사 범위와 재현 명령

**통합 9개:** 7구역 렌더, NPC/방문객 보안, 랙 앞·뒤/ToR/NIC/PDU·카세트, MMR 광 경로, 고객 A/B/C Co-location, MEP 전력·냉각·광, 모바일 두 손가락 조작, 서버실/Legacy 격리, 모바일 랙 제어 접기/펼치기.

**별도 5개:** 7구역 WebGL2 화면/보안, 랙 조작, MMR/Co-location/MEP, 모바일 화면/시설 격리, 실제 Chromium requestAnimationFrame 프레임 지표 수집. 실제 성능 측정치는 여기에서 인용하지 않음.

저장소 루트에서:

```bash
npm install --no-save --no-audit --no-fund @playwright/test@1.56.1 three@0.186.1
npx playwright install --with-deps chromium
python3 -m http.server 8765 --bind 127.0.0.1
# 다른 터미널:
npx playwright test --config DCI/DataCenter/ls3d-three/step19.playwright.config.mjs
# 별도 QA (자체 webServer 사용):
npx playwright test --config DCI/DataCenter/ls3d-three/playwright.step19.config.mjs
```

참고: 위 명령은 같은 8765 포트에서 웹 서버를 중복 실행하지 않도록 **선택적으로** 실행하세요. 첫 번째 테스트는 위 별도 로컬 서버를 사용하고, 두 번째 QA 설정에는 자체 `webServer`가 있습니다. 임시 Node 패키지 설치 방식은 기존 저장소의 설치 체계를 바꾸지 않습니다.

## 6. 반복 오류 이메일 재발 시 대응

1. 메일 제목의 **워크플로 이름/실행 링크/커밋 SHA**를 먼저 확인. 이전 실패(#49, #54~56, #10) 메일은 실행 기록이 남는 과거 사건이며 이후 성공으로 삭제되지 않음.
2. [전체 GitHub Actions](https://github.com/kimheeseo/LSCNS/actions)에서 새로운 실패인지 확인. Step 19의 **두 QA 이름을 구분**하여 확인.
3. 실행 로그에서 `intercepts pointer events`, `timeout`, `404`, `pageerror`, `root.visible` 등의 원인을 찾고 실제 UI가 실패하는지 검토. `force: true` 클릭 등으로 오류를 숨기지 않음.
4. 다른 Phase 1/2/4–8 워크플로의 과거 실패는 **이 보고서의 Step 19 두 워크플로 성공과 별개**입니다. 실패 메일이 계속된다면 해당 워크플로 이름을 기준으로 별도 확인 필요.
5. 메일 양을 줄이려면 GitHub의 Actions 알림 설정 또는 해당 워크플로의 `push.paths`를 **업무 중요도에 맞춰** 조정. 오류 원인을 확인하기 전에는 필수 QA 전체를 비활성화하지 않음.

## 7. 합격 범위 제한

이 검증은 GitHub-hosted Linux runner의 Chromium/SwiftShader 환경에서 수행됐습니다. 실제 현장 GPU 서버·케이블·전력·냉각·광 링크의 물리 시험을 의미하지 않습니다. 실제 Android/iOS 기기별 제스처, Safari 등 크로스브라우저, FPS/메모리의 장기 실측, 2D Gold Pixel Tour·BOM 회귀 전체 및 모든 과거 Phase 워크플로의 완전 합격은 이 결과만으로 보장하지 않습니다.

**문서 결론:** 반복 실패 원인을 수정하고, 해당 Step 19 통합 QA 9/9 및 별도 QA 5/5와 Pages #937 성공을 GitHub Actions 결과로 확인했습니다. 버전명을 임의로 승격하지 않고 기존 기능과의 호환 확인을 계속 권고합니다.
