# 데이터센터 비주얼 업그레이드 · 2026-11-26 KST

기준 커밋: `0b9099d17d7011a97cf689279842d3f39cf7e2ed`

## 반영 내용

- 실제 Three.js 장면의 네이비 배경·안개, 중성 작업 조명, 약한 따뜻한 보조광을 적용했습니다.
- 기존 물체 색과 metalness를 보존하고 roughness·공유 미세 범프·작은 절차적 반사 환경을 적용했습니다.
  외부 이미지 다운로드, 새로운 설비 메시, 새로운 렌더 루프는 없습니다.
- `#assetVisual`의 기존 SVG를 유지한 채 승인 이미지가 성공적으로 로딩된 경우에만 이미지 미리보기를 표시합니다.
- 시설 프로필·구성 서명·존·장비·부품 ID를 정확히 비교합니다. 구성 서명은 현재 모델에서 다시 읽어 시설 전환과 후속 확장의 변경도 반영합니다.
- 존 버튼·장비·부품 선택 후 같은 대상의 이미지가 연결됩니다. 이미지를 클릭하면 원래 모델의 이름·설명을 사용하는 확대 창이 열립니다.
- Escape/닫기/포커스 복귀를 지원합니다. 확대 창을 닫는 Escape는 기존 랙 검사·방문자 투어의 Escape 동작으로 전달하지 않습니다.
- BOM의 기존 노드 내부 빈 공간에만 이미지 슬롯을 추가했습니다. 노드·RU 위치·텍스트·연결선·경고색은 유지했습니다.
- 미리보기는 화면에 들어왔을 때 thumbnail/preview를 요청합니다. 4K 원본은 원본 링크를 눌렀을 때만 요청합니다.
- 서비스워커 버전을 갱신하고 이미지 매니페스트는 네트워크 우선으로 확인합니다.

## 이미지 상태

**Higgsfield에서 제작한 완성 이미지가 이번 작업에 제공되지 않았습니다.**
첨부된 릴스 스크린샷은 스타일 참고이며, 저장소의 기존 참고 이미지도 현재 시설 구성과 일치하는 새 이미지로 간주하지 않았습니다.

운영용 `assets/visual-hq/manifest.json`은 승인 이미지 0건입니다.
`pending-targets.json`의 실제 ID 24개는 제작 대기 목록이며, 존재하지 않는 이미지 URL을 요청하지 않습니다.
이미지 등록 방법·규격은 같은 폴더의 README에 있습니다.
팝업·실패 복귀·BOM 아이콘 검사는 미리보기 전용 PNG 시험 자료로 수행한 뒤 시험 자료를 모두 제거했습니다.

## 보존·검증 결과

| 검사 | 결과 |
|---|---|
| 기존 계산·신규 이미지·재질·랙 제어 관련 검사 | 61건 통과, 실패 0건 |
| 변경 전후 전체 계산 결과 | 12개 입력 시나리오 모두 동일 |
| 보호 파일 6개 | SHA-256 모두 동일 |
| 위 두 보존 검사의 재현 가능한 자동 검사 | 13건 통과, 실패 0건 |
| BOM 브라우저 비교 | KPI, BOM 표, 랙·냉각·전력 도식 문구, 광 링크 상세 모두 동일 |
| 데스크톱 레이아웃 | 주요 패널 5개 위치·크기 동일 |
| 장비 미리보기 | 기존 248×132px 유지, 모바일 표시 모드에서도 높이 132px 유지 |
| 이미지 확대 | 장비·존·BOM 아이콘 연결, 기존 이름·설명, 키보드 열기/닫기·포커스 복귀 확인 |
| 이미지 실패 | 잘못된 이미지의 디코딩 실패 후 기존 SVG 복귀 확인 |
| 랙 검사 간섭 | 이미지 창 Escape 후 랙 단독 보기 유지 확인 |
| 기본 URL / engine=legacy | Three.js 캔버스를 만들지 않는 기존 선택 방식 유지 |
| engine=three | 환경의 WebGL2 미지원에 따른 기존 대체 화면 전환 확인 |
| 페이지 자체 콘솔 오류 | 확인한 화면에서 0건. 브라우저 확장 프로그램 자체 오류는 별도 |
| 매니페스트 검사 | 승인 0건, 제작 대기 24건, 누락 이미지 요청 0건 |

보호 파일: `engineering-v82.js`, `calculation-definitions.js`, `capacity-model.js`,
`fiber-capacity.js`, `cable-power.js`, `LS_Datacenter_3D_Failure_Timeline_Model.js`.
시나리오와 전체 결과 해시는 `tests/assets/visual-hq-preservation-20261126.json`에 기록했습니다.

### 검증의 한계와 기존 실패

검증용 브라우저는 WebGL2를 지원하지 않습니다. 따라서 실제 Three.js 재질·조명 화면과 GPU를 사용하는
부품 이동·케이블 선택·실제 드래그 렌더링의 최종 외관 검수는 완료하지 못했습니다.
코드·재질·랙 제어 검사와 대체 화면의 선택·장애 시나리오는 확인했습니다.
모바일은 기존 툴의 모바일 표시 모드로 점검했으며 별도 실기기 검사는 하지 않았습니다.

전체 `ls3d-three/*.test.mjs`를 실행하면 106건 중 93건 통과·13건 실패입니다.
변경 전 소스에서도 해당 실패를 재현했습니다. 주요 항목은 기존 하역장 문구·드래그 계약,
네트워크 랙 RU·카세트·문·포트·배선 기대값, 누락된 외부 `three` 패키지,
기존 페이지 제목의 LS 접두사 기대값입니다. 이번 작업에서 이 기대값이나 모델 배치를 변경하지 않았습니다.
전체 3D 검사 통과 또는 실제 GPU 화면 검증 완료로 해석하면 안 됩니다.

## 주요 변경 파일

- `visual-hq-core.js`, `visual-hq.js`, `visual-hq.css`: 매칭·미리보기·팝업·지연 로딩.
- `assets/visual-hq/`: 빈 운영 매니페스트, 실제 ID 제작 대기 목록, 등록 안내.
- `ls3d-three/visual-finish.mjs`, `scene-adapter.mjs`: 표면·반사·조명.
- `LS_Datacenter_3D.html`, `ls3d-inspector-ui.js`, `ls3d-three/bootstrap.mjs`: 기존 갱신·선택 경로 연결과 캐시 버전.
- `index.html`, `app-v82.js`, `power-architecture.js`: 기존 BOM 노드 내 선택적 이미지 슬롯.
- `sw.js`: 캐시 버전·매니페스트 갱신.
- `scripts/verify-visual-hq.cjs`, 관련 테스트·보존 해시·CI: 이미지 등록 및 계산 보존 검사.

## 검사 명령

```sh
node DCI/DataCenter/scripts/verify-visual-hq.cjs
node --test DCI/DataCenter/tests/visual-hq-preservation.test.cjs DCI/DataCenter/tests/visual-hq.test.cjs DCI/DataCenter/ls3d-three/visual-finish.test.mjs DCI/DataCenter/ls3d-three/rack-hardware-detail.test.mjs DCI/DataCenter/tests/engineering-v82.test.cjs DCI/DataCenter/tests/calculation-validation.test.cjs DCI/DataCenter/tests/cable_power.test.cjs DCI/DataCenter/tests/fiber_capacity.test.cjs DCI/DataCenter/LS_Datacenter_3D_Failure_Timeline_Model.test.cjs
```

기존 engineering 검사는 결과 JSON을 다시 작성할 수 있으므로 검사 결과 파일을 코드 변경과 혼합해 커밋하지 않습니다.
