# 제품 카탈로그 검증 갱신 — 2026-11-17
## 범위와 한계
기존 499개 catalog.json / 21,882개 항목을 검사했다. 갱신 후 506개 catalog.json / 21896개 항목이다. **모든 항목의 최신 사양 확인을 완료한 것은 아니다.** 새 공식 자료로 확인한 항목과 미확인 RFQ를 분리했다. 공식 수치를 확인하지 못한 21,844개 항목은 표시/매칭용 specs를 null로 두고 legacy_specs에 원본을 보존했다. 숫자를 임의 추정하지 않았다. 대규모 RFQ 전환으로 구매 후보의 검증 수준이 크게 달라지므로 변경은 검토 가능한 PR로 제공한다.

## 기존 구조/스키마
- product_catalog/{업체}/{부품군}/[하위 제품군/Part NN/...]catalog.json: schemaVersion, company, category, displayName, checked, officialUrl, defaultSpecs, products.
- products는 ID/파일명을 키로 하는 사전이며 name, description, officialUrl/businessUrl/url, checked, sourceUrls, specs(한/영 혼합 키), referenceOnly/vendorReferenceOnly를 사용한다.
- catalog-manifest.json의 paths가 화면과 BOM 로더가 조회할 실제 파일 목록이다. bom-catalog-index.json은 역할별 보조 인덱스이며 주 로더는 paths를 직접 순회한다.
- product-catalog.js는 catalog.json을 카드/표로 표시한다. bom-catalog.js는 normalize→kind→match를 수행하고 bom-catalog-ui.js는 productRequirements를 후보와 연결한다.
- 기존 ID·경로·specs 키·API·수량 계산 엔진을 보존한다. 변경 전 specs/defaultSpecs는 legacy_specs/legacy_defaultSpecs로 보존한다. RFQ에는 모델명에서 속도·거리·폼팩터를 추론하는 기존 fallback을 적용하지 않는다.

## 추가 필드
source_url: 공식 개별 근거 URL 또는 null. 기존 officialUrl은 탐색 링크로 보존되며 검증 근거를 뜻하지 않는다.
verified_date: 실제 재검증 날짜 또는 null. 카탈로그 기준일과 구분한다.
spec_basis: 공식 datasheet / 추정 / null. 이번 작업에 추정 값은 없다. 공식 기술 매뉴얼/공식 제품 사양표도 공식 datasheet 범주에 포함한다.
status: active / EOL / announced / RFQ. RFQ는 사양 검증 상태이며 lifecycle_status가 별도로 있을 수 있다.
deprecated: 구형 세대 또는 판매 종료 표시. true 자체가 지원 종료/EOL을 뜻하지 않는다.
spec_sources: 현재 값이 있는 specs 키별 공식 URL. null 필드는 미확인/RFQ.
manufacturer/model/generation: 기존 업체/ID와 공식 확인 세대; 세대 미확인은 null.

## 변경 건수
|구분|추가|삭제|갱신|
|---|---:|---:|---:|
|제품 항목|14|0|21882|
|카탈로그 파일|7|0|499|

|현재 상태|항목 수|
|---|---:|
|active|17|
|announced|4|
|RFQ|21844|
|EOL(판매 종료)|31|
|deprecated 플래그(상태와 중복)|33|

정확한 파일별 추가/갱신 수: catalog-changes-20261117.json.
공식 사양 미확인 전체 목록: catalog-rfq-20261117.json (경로 → 제품 ID; 제품명과 빈 필드는 해당 catalog.json 참조).

## 카테고리 조사 결과
|분류|반영 모델/세대|검증 범위|
|---|---|---|
|GPU 서버|DGX B200/Blackwell, DGX B300/Blackwell Ultra|공식 GPU 수·폼팩터·시스템 전력. 기존 기본 B200 조건 유지|
|차세대 플랫폼|Vera Rubin NVL72, AMD MI455X/CDNA 5|공식 발표만으로 설계 SKU를 확정하지 않음; RFQ|
|스위치|Arista 7060X6, 7060XE7; NVIDIA QM9700/QM9790|공식 포트·속도·폼팩터, 모델별 전력 공개 시만 사용. XE7 출하 확인 전 announced|
|트랜시버|MMS4X00-NM16, MMA4Z00-NS400|공식 규격·거리·연마·프로토콜|
|AOC/DAC|MFA7U10-H00x, MCP4Y10-Nxxx|제품군 확인, 개별 길이/OPN 미확인 RFQ|
|광케이블/MPO|Corning EDGE G757512QPNDDU100F|공식 심수·OM4·길이. 극성은 null|
|PDU/UPS|APDU9941/NetShelter 9000, Galaxy VXL|공식 사양 문서, 미조회 값 null; 제품군 용량을 실제 설치 수량에 자동 적용하지 않음|
|냉각|Vertiv XDU1350|ATD 조건을 포함한 냉각용량과 펌프 운전별 전력|
|스토리지/NAS|NetApp AFF A90|공식 프로토콜; 용량·소비전력은 구성별 null|
|방화벽|FortiGate 700G|공식 데이터 인터페이스, 전력·RU null|
|라우터|Cisco 8223-64EF-M|공식 64×800G OSFP·3RU, 전력 null|
|LB|F5 r5900/rSeries|공식 L4/L7 수치, 전력·RU null|
|KVM|Raritan Dominion KX IV-101|공식 단일 포트, 전력·RU null|

## 주의할 공식 자료 차이
- MMS1X00-NS400: 현행 HTML 최대 500m, 과거 PDF ordering table 100m가 다르므로 거리 null/RFQ. 문서 revision/OPN 확인 필요.
- MMA4Z00-NS400은 SR4라도 MPO-12/APC이다. 일반 UPC 요구와 충돌 시 제외한다.
- F5 r5900은 공식 자료 L4 95Gbps를 사용했다. 검색된 비공식 100Gbps와 혼용하지 않는다.
- QM9700/QM9790: 64 logical NDR ports와 32 OSFP cages를 구분한다.
- 공식 Products No Longer For Sale 목록 31개는 EOL(판매 종료) + deprecated=true; 지원 종료 날짜는 추정하지 않는다.
- Arista XE7 공급 계획: 64PS/PRS Q4 2026, liquid/128PE Q1 2027. 분기 시작만으로 active로 전환하지 않는다.
- 기존 CPO 보조 인덱스의 Corning/Lumentum/NVIDIA 제품 수가 각각 1개씩 실제 데이터보다 작았다. 실제 개수로 재집계했다.

## 검증
{
  "gpu": 1024,
  "systems": 128,
  "bomLines": 32,
  "requirementRows": 32,
  "rowsWithCandidates": 26,
  "verifiedCandidateRows": 12,
  "errors": [],
  "sample": [
    {
      "item": "B200 8-GPU planning node",
      "n": 12,
      "first": "NVIDIA DGX B200",
      "status": "active"
    },
    {
      "item": "Compute NIC logical endpoints",
      "n": 9,
      "first": "NVIDIA BlueField-3 SuperNIC Family",
      "status": "RFQ"
    },
    {
      "item": "Leaf 400G planning switch",
      "n": 41,
      "first": "NVIDIA Quantum-2 QM9700",
      "status": "active"
    },
    {
      "item": "Spine 400G planning switch",
      "n": 41,
      "first": "NVIDIA Quantum-2 QM9700",
      "status": "active"
    },
    {
      "item": "server 400G-SR4 transceiver",
      "n": 216,
      "first": "NVIDIA MMA4Z00-NS400 transceiver",
      "status": "active"
    },
    {
      "item": "server 12F trunk 50 m",
      "n": 2292,
      "first": "Fiber Optic Data Center Trunk Cable",
      "status": "RFQ"
    }
  ]
}

- 기본 예제 BOM 32행 중 후보가 있는 행 26개; 공식 사양 확인 active 후보를 포함한 행 12개. 이는 전 품목 확정 매칭을 의미하지 않는다.
- RFQ는 참고 후보로만 표시되고 공식 사양 미확인 상태를 유지한다. EOL/announced는 구매 후보에서 제외한다.
- 계산 전후 수량 불변, RFQ 모델명 수치 추론 금지, EOL/announced 제외, 알려진 속도·커넥터·거리 충돌 검사를 수행했다.
- JS 구문 검사 및 JavaScript 격리 실행에서 엔진/매칭/검증 로직을 실행했다. Node CLI와 실제 브라우저 실행 환경이 없어 npm/Node 원본 명령 실행 및 렌더링 확인은 하지 못했다.
- 재현: node --test DCI/DataCenter/tests/catalog-audit-20261117.test.cjs DCI/DataCenter/tests/bom-catalog.test.cjs
- 기존 매칭 회귀 테스트는 archive를 명시적으로 읽는 과거 스키마 fixture로 수행한다. 현재 RFQ 제품을 검증 완료로 간주하는 검사가 아니다.

## 변경 파일
- product_catalog 내 기존 499개 catalog.json과 신규 7개 catalog.json
- product_catalog/catalog-manifest.json, bom-catalog-index.json, cpo-import.json, catalog-audit.json
- product_catalog/catalog-provenance.schema.json
- product_catalog/catalog-rfq-20261117.json, catalog-changes-20261117.json
- bom-catalog.js, bom-catalog-ui.js, product-catalog.js, index.html
- tests/catalog-audit-20261117.test.cjs, tests/bom-catalog.test.cjs
- CATALOG-UPDATE-20261117.ko.md (본 문서)
