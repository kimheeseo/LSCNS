# DataCenter Product Catalog

이 폴더는 DataCenter Tool의 **업체별 부품 리스트** 데이터 원본입니다.

## 폴더 규칙

```text
product_catalog/
├─ Corning/
│  ├─ EDGE8® Adapter Panels, LC/
│  │  ├─ catalog.json
│  │  ├─ EMOD8-CP08-AC_NAFTA_AEN.pdf
│  │  └─ ...
│  └─ ...
├─ SENKO/
└─ ...
```

- 1단계 폴더명 = 업체명
- 2단계 폴더명 = 부품군
- 부품군 폴더 안의 PDF = 제품 자료
- `catalog.json` = DataCenter Tool에 표시할 간략 스펙

PDF만 추가해도 제품 목록에는 자동으로 나타납니다. 간략 스펙까지 표시하려면 같은 폴더에 `catalog.json`을 추가합니다.

## catalog.json 예시

```json
{
  "schemaVersion": 1,
  "company": "Vendor",
  "category": "Product family",
  "description": "간단한 제품군 설명",
  "checked": "2026-10-04",
  "officialUrl": "https://vendor.example/product-family",
  "defaultSpecs": {
    "Connector": "LC",
    "Fiber count": "8F"
  },
  "products": {
    "MODEL-001": {
      "name": "MODEL-001",
      "specs": {
        "Color": "Blue"
      }
    }
  }
}
```

제품 key는 PDF 파일명에서 `.pdf`, `_NAFTA_AEN`, `_AEN`을 제거한 모델명을 사용하면 됩니다.

## 하위 제품군 폴더

부품군 아래에 제품군 폴더를 한 단계 이상 추가해도 DataCenter Tool이 재귀적으로 PDF를 탐색합니다.

```text
Corning/
└─ Adapter Panels/
   ├─ EDGE8® Adapter Panels, LC/
   │  ├─ catalog.json
   │  └─ *.pdf
   └─ EDGE8® Adapter Panels, MTP®/
      └─ *.pdf
```

따라서 **업체 → 부품군 → 제품군 → PDF** 구조도 지원하며, 화면에서는 업체와 부품군을 선택한 뒤 하위 제품군별 제품 카드가 표시됩니다.

## 현재 카탈로그 인벤토리 (2026-10-05)

이번 메타데이터 커밋 직전 `product_catalog`에는 PDF **1030개** / 구조화 `catalog.json` **69개**가 있습니다.

이번 업데이트에서 Sumitomo Electric 신규 10개 제품군과 US Conec MT Ferrules를 구조화했습니다. Sumitomo 첨부 PDF 10개는 전용 동기화 workflow가 공식 SEL 문서를 각 제품 폴더에 저장하며, 동기화 완료 후 카탈로그 audit를 다시 계산합니다.
