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

**현재 GitHub Pages 정적 배포 방식에서는 PDF 파일만 추가해도 제품 목록에 자동 표시되지 않습니다.** 폴더에 `catalog.json`을 작성한 뒤, `product_catalog/catalog-manifest.json`의 `paths` 배열에 해당 `catalog.json`의 상대경로를 추가해야 합니다. URL/제품 스펙은 `catalog.json`에서 관리합니다.

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

Sumitomo Electric은 첨부 10개 자료와 대응하는 공식 SEL PDF를 각 제품 폴더에 저장했고, **10개 제품군 / 54개 구조화 제품·구성 행**으로 정리했습니다. US Conec MTP® Connectors는 공식 제품 API 76페이지의 **1,895개 제품 행**을 drawing PDF 파일명 기준 **154개 제품 그룹**으로 묶어 추가했습니다. PDF 원본은 업로드하지 않고 공식 제품/도면 링크와 간략 스펙 메타데이터만 저장했습니다. NVIDIA는 Vera Rubin Platform에 이어 추가 자료 기준으로 **6개 제품군 / 27개 GPU·플랫폼·CPU·네트워킹·전문 GPU reference 제품**을 메타데이터 전용 catalog로 확장했습니다. Fujikura는 공식 optical-product 사이트의 hyperscale data center 솔루션 및 제품 분류 기준으로 **3개 제품군 / 37개 제품·솔루션 행**을 추가했습니다. Amphenol은 IT Datacom market catalog 기준으로 **3개 제품군 / 34개 제품 행**을 추가했습니다. ZTT의 IT Cabinet / Containerized Data Center / Micro Modular Data Center도 최신 공식 사양표 기준으로 갱신했습니다.

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

## CPU 업체 추가 · 2026-10-05

Intel Xeon 대표 4개, AMD EPYC 대표 3개, Qualcomm Centriq 과거 서버 제품 3개를 각 업체의 `CPU/catalog.json`에 등록했습니다. 공식 URL·코어·전력·메모리·확인 상태를 제공합니다. PDF 업로드는 하지 않습니다. Centriq의 현재 공급/지원은 검증 필요이며 신규 서버 BOM 자동 추천에서 제외됩니다. CPU TDP는 서버 전체 IT 부하와 별개입니다.

## 가속기·CPU 제품군·서버 인프라 확장 (2026-10-05)

Intel Gaudi 2/3와 Flex 140/170/170V(5개), Qualcomm AI 가속기 4종·연결 DSP 5종·Dragonfly C1000(10개), AMD EPYC 9006/9005/9004 제품군(3개), Supermicro(4개), HPE(4개), Dell(5개)를 추가했습니다. 총 신규 31개 항목, 신규 카탈로그 14개입니다. AMD 제품군은 기존 개별 CPU SKU와 별도 분류합니다. Qualcomm Centriq 과거 참고 3종은 보존합니다.

서버 인프라는 대표 제품을 등록하며 전체 SKU 목록을 의미하지 않습니다. 서버/랙 참고 수량과 케이블 BOM을 구분합니다. DSP 레인 속도와 전체 속도, 전원/냉각 용량과 실제 IT 부하를 구분합니다. 모든 제품은 공식 URL·확인일·근거와 검증 필요 사항을 포함하며 PDF 업로드는 없습니다.

## 트랜시버·광칩·스위치 확장 (2026-10-05)

- Coherent: 6개 항목
- AOI: 6개 항목
- Lumentum: 5개 항목
- Credo: 12개 항목
- Broadcom: 4개 항목
- MACOM: 5개 항목
- Marvell: 5개 항목
- Juniper: 4개 항목
- Cisco: 3개 항목
- Arista: 4개 항목

신규 업체 10개, 제품군 폴더 12개, 대표 제품/제품군 54개 항목입니다. Coherent/AOI는 트랜시버, Lumentum/Credo/Broadcom/MACOM/Marvell은 광칩·광원·PIC·DSP·TIA/드라이버, Juniper/Cisco/Arista는 데이터센터 Ethernet 스위치로 분류했습니다. 미공개 SKU와 공급/호환은 검증 필요로 표시합니다. 시연 제품·제품군은 referenceOnly로 구분합니다. 광칩을 완성 트랜시버 BOM에 중복 합산하지 않으며 자동 호환성 매핑은 추가하지 않았습니다. PDF 원본은 업로드하지 않습니다.


## 전력 설비·랙·배선 확장 (2026-10-05)

24개 대표 제품/제품군, 11개 신규 카탈로그입니다. UPS·발전기·랙·변압기·케이블 관리·Cat6A/DAC 동 케이블·AEC를 분류했습니다. Schneider Electric, Eaton, Cummins, Caterpillar, Hitachi Energy, Panduit 업체를 추가하고 기존 Amphenol/Credo에 배선 분류를 확장했습니다.

UPS 배터리/이중화, 발전기 DCC/Standby, 변압기 kVA·전압·보호 협조는 현장 설계 검증이 필요합니다. 랙 크기/하중은 IT kW가 아닙니다. AEC·DAC·ACC·AOC를 구분하고 광모듈 중복 합산을 피합니다. 카탈로그 등록이며 신규 설비의 자동 수량 계산/제품 호환 매핑은 포함하지 않습니다. PDF 업로드와 근거 없는 가격 추가는 없습니다.
