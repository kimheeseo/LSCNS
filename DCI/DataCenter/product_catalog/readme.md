# DataCenter Product Catalog

이 폴더는 DataCenter Tool의 **업체별 부품 리스트** 데이터 원본입니다.

## 폴더 규칙

```text
product_catalog/
├─ Corning/
├─ Furukawa Electric/
├─ SENKO/
├─ Sumitomo Electric/
├─ USConnec/
├─ YOFC/
└─ ZTT/
```

- 1단계 폴더명 = 업체명
- 2단계 폴더명 = 부품군
- 부품군 폴더 안의 PDF = 제품 자료
- `catalog.json` = DataCenter Tool에 표시할 제품/간략 스펙
- PDF 없이 `catalog.json`만 있어도 공식 URL 기반 제품 카드가 표시됩니다.

## catalog.json 예시

```json
{
  "schemaVersion": 1,
  "company": "Vendor",
  "category": "Product family",
  "description": "간단한 제품군 설명",
  "checked": "2026-10-05",
  "officialUrl": "https://vendor.example/product-family",
  "comparisonFields": ["Connector", "Fiber Count"],
  "products": {
    "MODEL-001": {
      "name": "MODEL-001",
      "description": "공식 제품 설명",
      "officialUrl": "https://vendor.example/model-001",
      "specs": {
        "Connector": "LC",
        "Fiber Count": "8F"
      }
    }
  }
}
```

## 현재 카탈로그 인벤토리 (2026-10-05)

현재 GitHub `product_catalog`에는 PDF가 총 **1050개** 있습니다. 모든 PDF는 업체별 부품 리스트의 재귀 탐색 대상입니다.

Sumitomo Electric은 현재 **30개 제품군 / 30개 공식 PDF**를 업체 → 부품군 → 제품군 구조로 분류했고, 각 폴더의 `catalog.json`에 홈페이지용 핵심 사양을 구조화했습니다.

US Conec MTP® Connectors는 공식 제품 API 76페이지의 **1,895개 제품 행**을 drawing PDF 파일명 기준 **154개 제품 그룹**으로 묶어 추가했습니다. PDF 원본은 업로드하지 않고 공식 제품/도면 링크와 간략 스펙 메타데이터만 저장했습니다.

SENKO는 Products 메뉴의 connector 계열을 기준으로 **5개 제품군 / 112개 구조화 제품 행**을 추가했습니다. 분류는 **VSFF, MPO/MT, SC/LC, Field, Legacy connectors**이며 PDF 없이 공식 제품 URL과 Store API description/간략 스펙만 저장했습니다.

ZTT의 IT Cabinet / Containerized Data Center / Micro Modular Data Center도 최신 공식 사양표 기준으로 갱신했습니다.

구조화 `catalog.json`이 아직 없는 PDF 제품군은 **34개 폴더 / 739개 PDF**입니다.
