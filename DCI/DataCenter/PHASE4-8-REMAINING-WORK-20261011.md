# LS Datacenter 3D — Phase 4~8 잔여업무 점검 (2026-10-11)

## 검토 대상
- 실행: https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html
- 소스: `LS_Datacenter_3D.html`, `LS_Datacenter_3D_Phase45.js`, `LS_Datacenter_3D_Phase6.js`, `LS_Datacenter_3D_Phase7.js`, `LS_Datacenter_3D_Phase8.js`
- 기존 가이드: `LS_Datacenter_3D_Phase8_Guide.md`

## 이번 변경 사항
- Phase 7: Ethernet+CPO처럼 서로 호환되지 않는 프로토콜/광 엔진 조합에서 스위치 측 플러거블 광모듈 수량을 **0으로 확정하지 않도록** 수정. 불명확한 조합은 `null/미산정` 유지.
- Phase 7: 감사 결과 JSON에 `compatible` 필드 추가. `blocking` 경고가 1건이라도 있으면 `false`; 단 `true`는 구매 적합성 승인이 아닌 **현재 구현된 차단 규칙의 미검출**임.
- 기존 HTML, 화면 탐색/렌더 파이프라인 및 기타 Phase 파일은 이 수정으로 변경하지 않음.

## 순차 점검 결과 및 후속 작업

### Phase 4 — GPU 캐시·LOD·컬링
- 현황: 캐시/프러스텀 컬링/LOD 상태와 캐시 ON/OFF UI가 존재함.
- 잔여: 정식 GPU Instancing 셰이더/버퍼, 반복 객체에 대한 CPU/GPU frame time 비교, 캐시 무효화 시 화면 정합성 자동 검사.
- 완료 기준: 같은 장면과 카메라에서 시각적 회귀가 없고, 캐시/Instancing ON/OFF 프레임 타임 비교 기록을 확보.

### Phase 5 — PC·Android·iPhone
- 현황: 각 기기에서 10초 rAF 측정 후 FPS/p95 및 JSON 추출 기능이 존재함.
- 잔여: Windows(Chrome/Edge, 하드웨어 GPU), Android(Chrome), iPhone(Safari) 실기기 결과.
- 완료 기준: 기기/브라우저/운영체제/모드/구역/품질별 JSON, 중앙값 및 p95, 발열 영향, WebGL 실패 로그. 데스크톱 헤드리스 결과를 모바일 실기기 테스트로 표기하지 말 것.

### Phase 6 — 냉각·전력 MEP
- 현황: 가상 Supply/Return 배관과 버스웨이, AABB 이격 검사 존재.
- 잔여: 실제 3D 형상 기반 상세 검사, 접속부/밸브/엘보/전원 분기와 유지보수 접근 공간, 케이블 최소 굽힘 반경.
- 완료 기준: 경로별 간섭/이격 실패 재현 케이스 및 정상 케이스 자동화, 각 검사의 입력 형상/단위/기준 공개. 현재 AABB로는 현장 설치 승인을 주장할 수 없음.

### Phase 7 — NVIDIA 광/BOM
- 현황: `product_catalog/NVIDIA/` JSON 로드, InfiniBand/Ethernet 선택, CPO/Pluggable 조건, 검토용 수량 산출, 기존 BOM 도구 열기 기능 존재.
- 이번 조치: Ethernet+CPO 시 스위치측 광모듈 0 확정 오류 방지 및 호환성 표시.
- 잔여: 공통 SKU/포트/링크 스키마를 정의하여 원본 BOM 엔진과 양방향 동기화, 카탈로그 모델별 포트·OSFP cage·breakout, 케이블/트랜시버 호환성 매트릭스 확인.
- 완료 기준: 3D↔BOM 변환 후 링크 수·거리·제품 ID·가격이 아닌 물량(BOM quantity) 일관성 검사 통과. SKU 미확정은 절대 구매 가능 `orderable=true`로 표시하지 않음.
- 필수 회귀 시나리오:
  1. IB+CPO: Q3450-LD 스위치측 pluggable=0, host-side 미산정.
  2. IB+Pluggable: switch-side 수량 미산정, cage/포트 일대일 대응 가정 금지.
  3. Ethernet+CPO: 차단 경고, switch-side pluggable=미산정.
  4. Ethernet 800G NIC 단일 포트: 차단 경고.
  5. 카탈로그 로드 실패: 숫자 0으로 대체하지 않고 오류 처리.
  6. `compatible=true`는 제조사 승인이나 구매 승인이 아님.

### Phase 8 — 통합 회귀/사용 가이드
- 현황: 브라우저 내부 12항목 API 스모크 검사, 사용자 가이드 및 Chromium 회귀 워크플로가 존재.
- 잔여: Phase 7 부정 조건 자동화, 성능 결과 증빙 보관 위치와 배포 URL/커밋 SHA 연결, 모바일 실기기 JSON 기준선.
- 완료 기준: 실패한 테스트는 결과 보고서에 명시하며, UI 컴포넌트 존재 검사를 업무 시나리오 성공으로 과대 해석하지 않음.

## 검증 상태 및 해석
| 검증 항목 | 상태 |
|---|---|
| GitHub 소스 확인 | 확인 |
| Phase 7 오산정 방지 코드 반영 | 반영 |
| Phase 7 코드 변경 후 자동 테스트 | 미실행 |
| 브라우저 Chromium 전체 회귀 재실행 | 미실행 |
| Windows/Android/iOS 실기기 | 미실행 |
| BIM/CAD 시공 정밀 검증 | 범위 밖 |
| 제조사 실제 호환성/구매 BOM 검증 | 미완료 |

**주의:** 과거 보고된 `78개 검사 통과/JS 오류 0`은 과거 실행 기록이며 이번 변경 사항의 신규 검증 결과가 아닙니다. 이후 워크플로 성공 기록/결과 JSON으로 재확인해야 합니다.

## 추가 구현 기록 (2026-10-11)
- Phase 4~5: 같은 브라우저 세션에서 측정한 완료된 FPS 기록을 최대 30건 보관하고, 기존 JSON 내보내기에 품질 설정/캐시 ON·OFF별 sessionHistory를 함께 담음. 실제 스마트폰 결과를 만들어내지 않음.
- Phase 6: 경로 전체 bounding prism 방식 대신 **선분(중심선)–AABB 최단거리에서 배관 반경을 뺀 값**으로 거리 계산 변경. 가상 랙 치수와 개념 좌표는 그대로임.
- Phase 8: 선분이 랙 위로 통과하는 비접촉 예(예상 2.9m)와 선분이 랙 내부를 관통하는 예(예상 0m) 두 회귀 항목 추가. 자동 워크플로의 최신 실행 결과는 별도 확인 필요.
- Phase 7: 앞선 수정대로 Ethernet+CPO의 플러거블 모듈 수량 확정을 금지. 본 변경에서는 기존 BOM 엔진의 양방향 계산 로직까지 교체하지 않았음.

### 남은 미완료 사항 (과장 금지)
1. Phase 4 실제 GPU instanced draw call 구현 및 GPU 시간 벤치마크는 미완료.
2. Phase 5 Windows GPU/Android/iPhone 실기기 측정 기록은 미확보.
3. Phase 6 상세 배관 CAD, 굽힘 반경, 서비스 접근성, 시공 안전 기준 검토는 미완료.
4. Phase 7 원본 BOM과 양방향 공유 스키마/엔진 동기화는 미완료.
5. Phase 8 신규 변경분의 Chromium CI 성공 여부는 아직 미확인. 문서상 회귀 항목 존재는 통과를 의미하지 않음.
