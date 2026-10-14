# STEP 06 — 엔지니어·경비원 3D 캐릭터 및 보행 애니메이션 개발 보고서

- 작성일: 2026-10-15 KST
- 저장소: kimheeseo/LSCNS (main)
- 판정: **코드 구현·구문 확인·Three.js API 모의 통합검증 17/17 PASS. 실제 브라우저 GPU 검증은 미실행.**
- 개발 범위: 20개 작업 중 06. **STEP 07 NavMesh/A* 충돌 회피 및 STEP 08 점검 상태머신은 착수하지 않음.**

## 1. 기존 대비 개선

기존 STEP 03의 단순 사각형 인물 모형을 Three.js Group 계층형 관절 캐릭터로 대체하였다. 작업자의 현재 좌표(x,z), 진행 방향(heading), 보행/대기/점검(action)은 기존 npcPose(n)에서 그대로 가져온다.

| 역할 | 외관 차별화 | 실제 Three.js 구성 |
|---|---|---|
| 엔지니어 2명 | 청록 작업복, 황색 안전모, 반사조끼 | 신체/팔/다리, 안전모·반사띠·사원증·무전기 |
| Utility 기술자 1명 | 보라톤 작업복, 황색 안전모 | 작업복·안전화·무전기·사원증 |
| OPS 운영자 1명 | 청색 복장과 헤드셋 | 모자·헤드셋·통신장비와 독립 팔다리 |
| 경비요원 2명 | 남색 제복, 경비모 및 표식 | security cap/visor, 사원증, 전용 복장 |

위 모델은 현장 교육/시각화용이며 실제 제조사 인증 보호구나 신체 물리학 모델이 아니다.

## 2. 관절 구조와 동작

캐릭터마다 다음 실제 Three.js.Group 피벗과 Mesh를 생성한다.

    npc:<id>
      npc-joint:torso
        npc-joint:neck-head
          npc-joint:helmet
        npc-joint:id-badge
        npc-joint:left-shoulder -> npc-joint:left-elbow
        npc-joint:right-shoulder -> npc-joint:right-elbow
      npc-joint:left-hip -> npc-joint:left-knee
      npc-joint:right-hip -> npc-joint:right-knee

- walk: 좌우 다리 교대 회전, 팔 역방향 스윙, 무릎 굽힘, 몸통 바운스.
- inspect: 왼팔 들어 올림/팔꿈치 회전, 입체 태블릿 모델 표시.
- pause: 팔 내림/미세 시선 및 호흡.
- idle: 경비원 정지, 미세 머리 움직임/호흡.

외부 애니메이션 타이머나 새 requestAnimationFrame을 만들지 않고 기존 렌더 루프 안에서 업데이트한다. 기존 npcPose() 경로를 변경하지 않으므로 STEP 07 NavMesh 또는 충돌회피가 구현되었다고 볼 수 없다.

## 3. 코드 및 연결

- [npc-characters.mjs](../ls3d-three/npc-characters.mjs): 신규 createNpcCharacterFactory(THREE,{cube,cylinder}). 역할별 3D PPE, ID카드, 무전기, 팔꿈치/무릎 관절과 상태별 움직임.
- [scene-adapter.mjs](../ls3d-three/scene-adapter.mjs): 옛 personMesh/refreshPeople 애니메이션을 교체. 기존 T.npcs, T.npcPose, facilityProfiles를 그대로 사용. 캐릭터 재사용/구역 제거/복귀와 진단 연결.
- [bootstrap.mjs](../ls3d-three/bootstrap.mjs)와 [LS_Datacenter_3D.html](../LS_Datacenter_3D.html): ESM 캐시 버전 step06 반영.
- [step06.playwright.spec.mjs](../ls3d-three/step06.playwright.spec.mjs): 6인 역할/안전모/경비모/ID/관절/점검 태블릿/회사 서버실·모바일/레거시 회귀 시험 작성(실기기 미실행).

캐릭터는 기존 공유 BoxGeometry/CylinderGeometry와 PBR Material을 재사용하고, 매 프레임 새로운 Geometry를 만들지 않는다. 제거된 캐릭터는 레지스트리에서 해제하고, 필요 시 재생성한다.

**기존 앱 비변경 요소:** NPC 원래 경로 및 시설 모드 데이터, 설비 ID, 랙 문·슬라이드, 시뮬레이션 시간/시나리오, BOM 및 전력·냉각 수식, 원본 2D Tour.

## 4. 검증 근거

### 4-A. 실제 GitHub 소스를 사용한 JS API Mock 통합 실행: 17 / 17 PASS

검증에 연결된 최신 소스는 npc-characters.mjs, scene-adapter.mjs, architecture.mjs, equipment-factory.mjs이다. 모의 THREE.Group/Mesh/Renderer 등으로 통합 프레임을 실행했다. 아래 PASS는 **GPU 화면 실제 픽셀과 FPS의 증거가 아니다**.

| 테스트 | 결과 |
|---|---|
| 캠퍼스 총 6인 생성 | PASS |
| 엔지니어 2/기술자 1/운영자 1/경비원 2 | PASS |
| 독립 캐릭터 그룹 | PASS |
| 머리·어깨·팔꿈치·무릎 피벗 | PASS |
| 안전모 모델 | PASS |
| 경비모 모델 | PASS |
| ID 카드 | PASS |
| 걸을 때 팔·다리 교대 운동 | PASS |
| 점검 자세 및 태블릿 표시 | PASS |
| 대기 보안요원 태블릿 숨김 | PASS |
| 반복 렌더 프레임의 객체 재사용 | PASS |
| Data Hall에서 eng-hall만 표시 | PASS |
| 회사 서버실 경비원 제거 | PASS |
| 상세 검사 시 NPC 숨김 | PASS |
| Campus 복귀 시 6명 복원 | PASS |
| 구역별 필요 캐릭터 재생성 | PASS |
| dispose() 정상 | PASS |

모의 테스트에서 초기 캠퍼스 NPC 6개, 모드/구역 왕복 후 누적 생성 12개를 기록했다. 누적 생성량은 FPS·GPU 리소스 실측값이 아니다.

### 4-B. 브라우저 회귀: 미실행

실제 WebGL2 브라우저 화면에서 안전모·관절 표현의 품질, 보행 중 끊김/바닥 관통, GPU/메모리/FPS, PC/모바일 오류를 독립 검증하지 못했다. 사용자에게 정상 작동으로 확정 보고하지 않는다.

브라우저 수동 검증 체크:
1. [Three.js 최신 실행](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three)에서 강력 새로고침.
2. Campus 전경에서 직원 4명과 경비원 2명의 복장·ID카드·안전모/경비모 식별.
3. Data Hall/Utility/NOC에 진입하여 각 역할의 보행·정지·점검 자세 확인.
4. 회사 서버실에서 외부 경비원 등 다른 구역 인물이 보이지 않는지 확인.
5. 기존 작업자 따라가기/1인칭 시점, 랙 문·서버 인출, 레거시 모드 정상 확인.
6. F12 Console에서 window.LS3D_THREE.adapter.npcAnimation 확인.
7. 실제 Playwright + WebGL2 검사 수행 시 오류 로그·스크린샷을 근거와 함께 첨부.

Playwright 사용법:

    npm install --no-save @playwright/test
    npx playwright install chromium
    npx playwright test DCI/DataCenter/ls3d-three/step06.playwright.spec.mjs

## 5. 남은 과제와 다음 업무

- STEP 06: 코드/정적/Mock 17/17 통과, 실브라우저 수동·자동 및 FPS 측정 대기.
- STEP 07: NavMesh/A* 경로 탐색, 문/랙/사람 충돌회피 (미착수).
- STEP 08: 설비 점검의 정식 상태머신과 장애 대응 (미착수).
- STEP 09: 출입접수·X-ray 보안 시퀀스 (미착수).

**현재 검증 수준을 실제 GPU 합격과 혼동해서는 안 된다. 7번 개발은 사용자 별도 승인 전 시작하지 않는다.**
