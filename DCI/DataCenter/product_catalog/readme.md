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

현재 총 **311개 PDF**가 등록되어 있으며, DataCenter Tool은 모든 하위 제품군을 재귀적으로 읽습니다.

- `Corning/Accessories/EDGE™ Solutions_Rack Accessory`: 8 PDFs
- `Corning/Accessories/MTP PRO Accessories`: 6 PDFs
- `Corning/Accessories/Reverse Polarity LC UniBoot, Duplex Clip`: 10 PDFs
- `Corning/Bracket/EDGE™ Strain-Relief Bracket`: 4 PDFs
- `Corning/Connector/MMC`: 1 PDFs
- `Corning/Housing/EDGE™ Housing, FX`: 5 PDFs
- `Corning/Housing/Pretium® Connector Housing (PCH)`: 5 PDFs
- `Corning/Jumper/12-Fiber MTP® PRO Jumper`: 18 PDFs
- `Corning/Jumper/EDGE™ Solutions Jumper, 2 F, LC Uniboot to SC Duplex`: 2 PDFs
- `Corning/ODF/EDGE™ Solutions_ODF`: 5 PDFs
- `Corning/Panel/Adapter Panels/EDGE8® Adapter Panels, LC`: 12 PDFs
- `Corning/Panel/Adapter Panels/EDGE8® Adapter Panels, MTP®`: 52 PDFs
- `Corning/Panel/Adapter Panels/EDGE™ Adapter Panel, MTP®`: 8 PDFs
- `Corning/Panel/CCH Panel/CCH Panel, FC Adapters`: 6 PDFs
- `Corning/Panel/CCH Panel/CCH Panel, LC Adapters`: 29 PDFs
- `Corning/Panel/CCH Panel/CCH Panel, MTP Adapters`: 20 PDFs
- `Corning/Panel/EDGE™ Panels`: 3 PDFs
- `Corning/Trunk/EDGE™ Armored Trunk`: 17 PDFs
- `Corning/Trunk/EDGE™ MTP® Extender Trunk`: 30 PDFs
- `USConnec/Fiber Optic Cleaners`: 41 PDFs
- `USConnec/MDC Connectors`: 6 PDFs
- `USConnec/MMC Connectors`: 23 PDFs

같은 제품군은 `comparisonFields`에 정의한 동일한 스펙 항목/순서로 카드형과 표형에서 비교합니다.
