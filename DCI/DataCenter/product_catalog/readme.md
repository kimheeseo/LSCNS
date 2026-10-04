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

현재 GitHub `product_catalog`에는 PDF가 총 **1021개** 있습니다. 모든 PDF는 업체별 부품 리스트의 재귀 탐색 대상입니다.

Sumitomo Electric은 첨부 10개 자료와 대응하는 공식 SEL PDF를 각 제품 폴더에 저장했고, **10개 제품군 / 54개 구조화 제품·구성 행**으로 정리했습니다. ZTT의 IT Cabinet / Containerized Data Center / Micro Modular Data Center도 최신 공식 사양표 기준으로 갱신했습니다.

구조화 `catalog.json`이 아직 없는 PDF 제품군은 **34개 폴더 / 739개 PDF**입니다.

### 구조화 메타데이터 재검토 필요
- `Corning/Accessories/EDGE™ Port Replication Housing Accessory`: 3 PDFs
- `Corning/Accessories/EDGE™ Solutions_Rack Accessory`: 8 PDFs
- `Corning/Accessories/MTP PRO Accessories`: 6 PDFs
- `Corning/Accessories/Reverse Polarity LC UniBoot, Duplex Clip`: 10 PDFs
- `Corning/Cable/Loose Tube/ALTOS® Figure-8 Loose Tube, Gel-Free Cable`: 62 PDFs
- `Corning/Cable/Loose Tube/ALTOS® HD Gel-Free, All-Dielectric Cable with Binderless FastAccess® Technology`: 8 PDFs
- `Corning/Cable/Loose Tube/ALTOS® HD Lite, Gel-Free, Single-Jacket, Single-Armored Cable with FastAccess® Technology`: 9 PDFs
- `Corning/Cable/Loose Tube/ALTOS® Lite Loose Tube, Gel-Free, Single-Jacket, Single-Armored Cables with FastAccess® Technology`: 10 PDFs
- `Corning/Cable/Loose Tube/ALTOS® Loose Tube, Gel-Free, All-Dielectric Cable with FastAccess® Technology`: 60 PDFs
- `Corning/Cable/Loose Tube/SOLO® ADSS Medium-Span, Loose Tube, Gel-Filled Cable`: 44 PDFs
- `Corning/Cable/Loose Tube/SOLO® ADSS Short-Span, Loose Tube, Gel-Filled Cable`: 42 PDFs
- `Corning/Cable/Outdoor Duct Cables/ALTOS® Loose Tube, Gel-Free Cable`: 149 PDFs
- `Corning/Connector/MMC`: 1 PDFs
- `Corning/Harness/EDGE™ Conversion Harness`: 3 PDFs
- `Corning/Harness/EDGE™ Non-Staggered MTP to LC Harness`: 44 PDFs
- `Corning/Harness/EDGE™ Solutions 24 F Y Harness`: 11 PDFs
- `Corning/Harness/EDGE™ Solutions 2x3 Conversion Harness`: 1 PDFs
- `Corning/Harness/EDGE™ Staggered Harness`: 29 PDFs
- `Corning/Module/EDGE™ 4x4 Mesh Module`: 2 PDFs
- `Corning/Module/EDGE™ Base-8 Module`: 2 PDFs
- `Corning/Module/EDGE™ Bidi Tap Module`: 2 PDFs
- `Corning/Module/EDGE™ Conversion Modules`: 4 PDFs
- `Corning/Module/EDGE™ Module, Ultra Low Loss`: 9 PDFs
- `Corning/Module/EDGE™ Module`: 11 PDFs
- `Corning/Module/EDGE™ Tap Module`: 22 PDFs
- `Corning/Module/Fiber to the Desk Module`: 8 PDFs
- `Corning/ODF/EDGE™ Solutions_ODF`: 5 PDFs
- `Corning/Panel/Adapter Panels/EDGE™ Adapter Panel, MTP®`: 8 PDFs
- `Corning/Panel/CCH Panel/CCH Panel, FC Adapters`: 6 PDFs
- `Corning/Panel/CCH Panel/CCH Panel, SC Adapters`: 30 PDFs
- `Corning/Panel/EDGE™ Panels`: 3 PDFs
- `Corning/Trunk/EDGE™ Hybrid Trunk`: 21 PDFs
- `Corning/Trunk/EDGE™ Indoor Ribbon Trunk`: 1 PDFs
- `Corning/Trunk/EDGE™ MTP® Trunk`: 105 PDFs
