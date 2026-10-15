# STEP 20 / 20 — 최종 릴리스 후보·문서화·GitHub 업로드 결과

- 기준일: **2026-10-15 KST**
- 저장소: **kimheeseo/LSCNS**, 브랜치 **main**
- 기존 베이스 **v4.7.0**, Three.js 20단계 결과 **v4.8.0-rc.1**
- 검증 판정: **20단계 구현 자료/README/TXT/정적 확인 완료. 실제 Chromium/WebGL2·GitHub Pages 배포·실측 성능 전체 PASS는 확인되지 않아 정식 v4.8.0 확정 보류.**
- 페이지: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three
- Legacy 복귀: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=legacy

## 1. 이 단계에서 실제 반영한 사항

| 요구 | 조치 | 상태 |
|---|---|---|
| 20단계 최종 작업 | 이전 01~19단계 결과·실브라우저 QA 필요 조건 정리 | 진행표/보고서 반영 |
| GitHub 업로드 | 저장소의 main에서 HTML 표시 버전·사용 설명서·버전 TXT·최종 보고서 변경 | 코드/문서 업로드 |
| 사용 방법 수정 | 저장소 루트 README 및 DataCenter/README.md를 직접 실행·상세 조작 중심으로 개선 | 완료 |
| 별도 버전업 TXT | DataCenter/LS3D_VERSION_UPGRADE_HISTORY.txt 신규 작성 | 완료 |
| 버전 표기 | 기존 v4.7.0 코어, 신규 v4.8.0-rc.1 Three.js 릴리스 후보 표시 | HTML title/사이드바 |
| 최종 검증 | 보존 파일 SHA·3D 엔진 링크/버전·모듈 구조·기존 소스 Mock 근거 확인 | 정적 검사 및 과거 Mock 검증, 실브라우저 PASS는 미확인 |

**GitHub Tags/Releases 객체는 만들지 않았습니다.** v4.8.0-rc.1은 HTML와 문서의 릴리스 후보 문자열입니다. 배포 전 종합 브라우저 PASS가 확인되지 않았으므로 정식 v4.8.0으로 판정하지 않습니다.

## 2. 개발 연혁

단계 01~03: 기존 코드/설계·Three.js r186.1 WebGL2/PBR 렌더 분기 도입.
단계 04~07: 건축 및 설비 3D/PBR/LOD, NPC 관절 애니메이션과 A* 경로.
단계 08~10: NPC 점검 FSM, 방문객 출입 보안, 랙 42/48U 부품.
단계 11~14: 전/후면 문, 개별 서버 인출/NIC, 광 카세트, ToR/전력 A/B.
단계 15~17: MMR 4구획, 고객 A/B/C Co-location, 광·전력·냉각 입체 연결.
단계 18~19: 카메라 가림 복원, 모바일 제스처/적응형 DPR, Node/Playwright 통합 QA 코드와 Actions.
단계 20: 버전 후보 표시, 사용 설명서 및 업그레이드 TXT, 최종 상태 기록.

세부 기술 내용/파일 경로/캐시 버전업 규칙은 **[버전업 TXT](../LS3D_VERSION_UPGRADE_HISTORY.txt)**에 기록했습니다.

## 3. 주요 링크 및 접근법

- **[README — 사용법](../README.md)**
- **[버전업 TXT](../LS3D_VERSION_UPGRADE_HISTORY.txt)**
- [저장소 첫 화면](https://github.com/kimheeseo/LSCNS)
- [20단계 개발 로드맵](../LS3D_V4_20_TASK_ROADMAP.md)
- [세부 검증 진행표](../LS3D_PROGRESS_STATUS.md)
- [19단계 통합 QA 보고서](./STEP19_INTEGRATION_QA_REPORT.md)
- [19단계 통합 Actions (8개 브라우저 시나리오)](https://github.com/kimheeseo/LSCNS/actions/workflows/ls3d-step19-quality.yml)
- [19단계 별도 Actions (5개 브라우저 시나리오)](https://github.com/kimheeseo/LSCNS/actions/workflows/ls3d-step19-qa.yml)

### 사용자 수동 사용 확인 절차

1. Three.js 링크를 열어 **Three.js r186 · PBR** 표시 및 3D 객체 여부 확인.
2. Campus 및 나머지 6개 구역 모두 진입, 카메라 회전/확대/설비 클릭.
3. Data Hall에서 랙 앞/뒤 카메라, 문, 개별 서버/NIC, 광 카세트, ToR 및 PDU A/B 연결.
4. Network/MMR에서 Carrier/ODF/DWDM/Core 구획과 각 출입문, 선택 광회선.
5. Data Hall 고객 A/B/C 구획과 광·전력 A/B 연결, Campus MEP 전력·냉각·광 overlay.
6. 모바일 폭 390px 및 두 손가락 pinch/pan, 화면에서 겹친 패널이나 오클릭 없는지 확인.
7. 회사 서버실 모드에서 외부 건물/경로 격리, Legacy 모드 전환/복귀.
8. 개발자 Console에서 JS exception, 404, MIME/CORS, WebGL 경고 확인.

이 절차는 **검증 지침**이며 직접 완료한 테스트 로그를 뜻하지 않습니다.

## 4. 확인된 검증과 아직 미확인 사항

### 확인된 근거

- STEP 03 사용자 수동 브라우저 기본 기능 **6/6 PASS** (과거 사용자 보고; 20단계 새 빌드 전체 합격과 구별).
- STEP 19에서 기존 STEP 14~18의 실제 GitHub 모듈을 JS/Three-compatible Mock 환경에서 재검증하여 **22/22 PASS**.
- STEP 19에서 핵심 JS 파일 20개 구문 검사 정상/확인한 상대경로 import 누락 0건.
- STEP 19에 실제 Chromium/WebGL2 수용 테스트 및 GitHub Actions YAML 2종 업로드.
- 현재 릴리스의 HTML 표시 버전 및 원본 보호 링크/파일이 main에 있는지 확인.
- 기존 2D Gold Pixel Tour/BOM의 SHA 및 변경 부재:

  - DCI/DataCenter/LS_Datacenter_Campus.html blob **9c84f24f2101043d7c113ce2d0da50e7245bc318**
  - DCI/DataCenter/index.html blob **692767f0c234df3b8310642e3acd032137c68baa**

### 미확인/보류

- 이 세션에서 공개 GitHub Pages와 GitHub Actions 최신 실행 결과를 웹으로 열 수 없어 **원격 배포 결과를 직접 확인하지 못함**.
- Chromium/WebGL2 화면 렌더·픽셀·터치·전체화면·카메라 가림·7구역 브라우저 통합 테스트 **PASS 로그/아티팩트 미확보**.
- 실제 FPS, GPU/메모리 사용량 개선 수치 및 환경별 성능·호환성 검증 미확보.
- 정상 동작이 확인되어야 **v4.8.0** 정식 릴리스로 승격 가능. 검증하지 않은 성능 수치는 보고하지 않는다.

## 5. 다음 정식 릴리스 승격 조건

1. GitHub Actions의 소스 테스트와 Chromium WebGL2 테스트를 실제 실행하고 각각 PASS·로그·화면 캡처를 확인한다.
2. 실패가 나오면 해당 stack trace, 3D 화면 캡처 및 자산 ID를 기준으로 P0/P1 수정하고 해당 테스트를 재실행한다.
3. GitHub Pages 배포본을 Chrome/Edge·390px 모바일에서 실제 수동 검증한다.
4. 확인된 평균/백분위 프레임 시간, GPU memory 정보는 측정 환경과 함께 기록한다.
5. 신규 캐시 Import query가 필요한 실행 코드 변경이라면 단계별 module URL 쿼리도 변경하고 Node/Playwright 재검증한다.
6. 단계를 만족하면 HTML/README/TXT 버전을 **v4.8.0**으로 승격하고 필요한 경우 별도 git tag / GitHub Release를 생성한다.

## 6. 운영 롤백 및 유지보수

- 긴급 사용자 우회: **?engine=legacy** 접속.
- 장애 추적: 브라우저 F12 Console/Network + GitHub Actions 로그 및 업로드된 아티팩트를 확인한다.
- 마지막 정상 커밋으로 복귀할 경우 HTML/Bootstrap/scene-adapter ES import URL의 query version도 함께 검토한다.
- 원본 2D, BOM 데이터 모델, lsdc-twin-bom/1.0, COMSOL/CPO 등 별도 도구 코드는 이 작업에서 변경하지 않는다.
- 본 모델의 경로, U/포트 및 시각적 배관은 교육/설계 검토용. 실제 CAD/BIM·IEC/TIA 적합성·전력/냉각 용량·광 링크 정격의 실측 검증 대체물이 아니다.

**결론:** STEP 20의 버전업 표기·사용법 정리·TXT 릴리스 이력·GitHub 소스 반영은 완료. 실 브라우저/배포 종합 수용성은 아직 완료라 판정할 수 없으므로 **v4.8.0-rc.1** 유지.
