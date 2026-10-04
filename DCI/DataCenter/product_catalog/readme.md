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

## 현재 카탈로그 인벤토리 (2026-10-04)

현재 GitHub에 업로드된 PDF는 총 **394개**입니다. 모든 PDF는 DataCenter Tool의 재귀 탐색으로 목록에 노출됩니다.

구조화된 `catalog.json`이 아직 없는 제품군은 **9개 폴더 / 125개 PDF**이며, 이 파일들은 PDF 목록은 표시되지만 동일 스펙 비교는 아직 미등록 상태입니다.

### 구조화 메타데이터 재검토 필요
- `Corning/Accessories/EDGE™ Solutions_Rack Accessory`: 8 PDFs
- `Corning/Accessories/MTP PRO Accessories`: 6 PDFs
- `Corning/Accessories/Reverse Polarity LC UniBoot, Duplex Clip`: 10 PDFs
- `Corning/Connector/MMC`: 1 PDFs
- `Corning/ODF/EDGE™ Solutions_ODF`: 5 PDFs
- `Corning/Panel/Adapter Panels/EDGE™ Adapter Panel, MTP®`: 8 PDFs
- `Corning/Panel/CCH Panel/CCH Panel, FC Adapters`: 6 PDFs
- `Corning/Panel/EDGE™ Panels`: 3 PDFs
- `Corning/Trunk/EDGE™ MTP® Trunk`: 78 PDFs

YOFC는 PDF 없이 공식 URL 기반으로 7개 페이지를 등록합니다. Product Model이 공개된 페이지는 모델별로 분리하고, Product Model이 없는 페이지는 Characteristics-only 항목으로 저장합니다.
