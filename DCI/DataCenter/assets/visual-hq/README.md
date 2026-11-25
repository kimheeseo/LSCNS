# 고화질 이미지 에셋 적용

현재 Higgsfield 완성 이미지가 제공되지 않아 `manifest.json`의 `assets`는 비어 있습니다.
`pending-targets.json`은 실제 시설 프로필·존·설비 ID에 대응하는 제작 대기 목록이며 브라우저에서 로딩하지 않습니다.
기존 참고 이미지를 다른 구성의 정확한 설계 이미지로 임의 연결하지 않았습니다.

## 이미지 등록

1. 승인된 PNG/WebP와 축소본을 이 폴더의 `media/` 아래에 저장합니다.
2. 대상 시설·존·장비·부품을 실제 툴에서 선택한 다음 개발자 콘솔에서
   `DCVisualHQ.captureContext()`를 읽습니다. 반환된 ID와 `composition`을 복사합니다.
   부품은 3D에서 선택된 실제 `partId`를 사용합니다. 존 이미지는 존 버튼을 선택한 상태에서 기록합니다.
3. 아래 형식으로 `manifest.json`에 등록합니다. 이름·설명·용량을 이미지 메타데이터에 만들지 않습니다.
4. `node DCI/DataCenter/scripts/verify-visual-hq.cjs`를 실행한 뒤 이미지·매니페스트를 함께 커밋합니다.
5. 실화면에서 다른 시설·장비·부품으로 전환해 잘못 재사용되지 않는지 확인합니다.

```json
{
  "scope": "tour",
  "status": "ready",
  "profile": "aidc",
  "zoneId": "hall",
  "assetId": "tor-a",
  "partId": null,
  "composition": "실제 captureContext()에서 복사",
  "view": "front",
  "width": 3840,
  "height": 2160,
  "alpha": false,
  "thumbnail": {"src": "media/dc-aidc-hall-tor-a-front-v01-480.webp", "width": 480, "height": 270},
  "preview": {"src": "media/dc-aidc-hall-tor-a-front-v01-1600.webp", "width": 1600, "height": 900},
  "full": {"src": "media/dc-aidc-hall-tor-a-front-v01-3840.webp", "width": 3840, "height": 2160}
}
```

존 이미지는 `assetId`와 `partId`를 `null`로 둡니다. 부품 이미지는 선택된 실제 ID를 모두 기록합니다.
프로필·구성·존·장비·부품은 모두 정확히 일치해야 합니다. 부모 이미지나 다른 랙 이미지로 대체하지 않습니다.
시설 구성 또는 렌더링 모델을 변경하면 `visual-hq-core.js`의 `REVISION`을 갱신하고 해당 이미지를 다시 검수합니다.
문·카세트의 실시간 위치는 이미지에 반영되지 않으며, 확대 창에 기준 시점 이미지라는 설명을 표시합니다.

## BOM 아이콘

BOM 화면에서 `DCVisualHQ.captureContext()`를 읽어 `profile`과 `composition`을 기록합니다.
`scope`는 `bom`, `view`는 `icon`, `alpha`는 `true`로 지정하고 tour의 세 ID는 `null`로 둡니다.
`diagram`에는 실제 슬롯의 `svgId`와 `nodeKey`를 사용합니다.
슬롯은 DOM의 `[data-hq-slot]`에서 확인하며 새 노드나 연결선은 만들지 않습니다.

- `rackSvg`: 기존 컴포넌트의 `label + '|' + startU + '-' + endU`.
- `coolingSvg`: 현재 표시된 `CRAH / HVAC`, `CDU`, `CHILLER PLANT` 중 해당 키.
- `powerArchitectureSvg`: 현재 노드 제목. 예: `UPS A`, `Transformer A`.

사용자 입력이 변경되면 BOM 구성 서명이 달라져 해당 입력에서 검수하지 않은 아이콘은 자동으로 원래 도식으로 돌아갑니다.

## 웹 규격

원본은 긴 변 3840px 이상을 권장하며, 화면비를 유지합니다. 투명 원본은 알파 PNG, 웹 제공본은 알파 WebP를 사용할 수 있습니다.
일반 팝업은 1600px 180~400KB, 썸네일은 480px 20~60KB, 아이콘은 256/512px 20~80KB를 목표로 합니다.
미리보기에는 thumbnail/preview만 사용하며 full은 사용자가 원본 링크를 누를 때 요청합니다.
페이지 진입 시 모든 4K 파일을 내려받거나 서비스워커 CORE에 넣지 않습니다.

이미지가 없거나 매칭되지 않거나 디코딩에 실패하면 기존 SVG를 유지합니다.
계산·도식 라벨·A/B 상태·부품 검사 UI·3D 캔버스는 이미지로 대체하지 않습니다.
