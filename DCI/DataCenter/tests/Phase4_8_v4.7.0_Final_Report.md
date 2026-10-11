# LS Datacenter 3D v4.7.0 · Phase 4–8 통합 검증 결과

**웹:** https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?v=4.7.0

**소스:** https://github.com/kimheeseo/LSCNS/tree/main/DCI/DataCenter

## 1. 단계별 구현 상태

| Phase | 구현 범위 | 검증·제약 |
|---|---|---|
| 4 | 정적 GPU 버퍼 캐싱, LOD, Frustum Culling, 라우트 구간 컬링 유지. 화면 품질/캐시 전환과 성능 진단 패널 추가 | 실제 셰이더 기반 GPU Instancing은 미구현 |
| 5 | 현재 브라우저의 WebGL/렌더러/해상도 정보, 10초 rAF FPS, p95 프레임 시간, JSON 증거 저장 | 실제 Windows/Android/iPhone 하드웨어 원격 검증 불가 |
| 6 | Data Hall 배관 공급·환수, 상부 Busway, 랙 케이블 경로 3D 개념도, Z 오프셋과 AABB 이격 검사, CSV | 가상 좌표·이격 사전검토, 실측 3D BIM/시공 승인 아님 |
| 7 | NVIDIA 제품 카탈로그를 직접 읽어 스위치·NIC·광 연결 BOM 초안, CPO와 OSFP cage 차이, 프로토콜 차단, 기존 3D→BOM 공유 기능 연결 | 실제 주문 가능 SKU, 포트 breakout, 광모듈 호환, 광예산 등 미확정 |
| 8 | 12항목 비파괴적 자가 점검, 종합 사용자 가이드, 기존 62항목 및 신규 PC/모바일 회귀 검증 | 실제 실기기·시공·구매 적합성 검증은 별도 |

## 2. NVIDIA 공식 제품 자료 기반 감사

- Q3450-LD CPO: InfiniBand XDR, 144×800G 논리 포트, 전면 144 MPO12. 스위치 측 외장 플러거블 광모듈 **0**.
- Q3400-RA: 144 논리 포트, 72 OSFP cage. 논리 포트 1개를 광모듈 1개로 확정하지 않으며 미산정 처리.
- ConnectX-8 SuperNIC: InfiniBand 1×800G 또는 2×400G; Ethernet 한 포트 800G를 지원한다고 일반화할 수 없음.
- Spectrum-4 SN5000: Ethernet 제품군. 정확한 SKU·포트 구성 전 구매 수량 미산정.
- 예시: 20 racks × 2 optical links/rack × 34m = **40 링크, 1,360m**(포설 여유 제외).
- 기존 BOM 도구 연동은 기존 3D 공유 스키마의 링크 열기 기능을 사용합니다. Phase 7의 감사 수량을 BOM 엔진 수량에 덮어쓰지 않습니다.

제품 데이터 원본: product_catalog/NVIDIA/InfiniBand Switch Models/catalog.json 및 product_catalog/NVIDIA/Networking/catalog.json. 공식 NVIDIA 제품 페이지/데이터시트는 각 카탈로그 레코드에 보존돼 있습니다.

## 3. Chromium 실행 결과

| 자동화 | PASS | FAIL | JS/콘솔 오류 | GitHub Actions |
|---|---:|---:|---:|---|
| 기존 Phase 1 FOV/랙/작업자 | 46 | 0 | 0 | https://github.com/kimheeseo/LSCNS/actions/runs/38112037623 |
| 기존 Phase 2 MEP/광 시나리오 | 62 | 0 | 0 | https://github.com/kimheeseo/LSCNS/actions/runs/38112037634 |
| Phase 4–8 통합 테스트 | 78 | 0 | 0 | https://github.com/kimheeseo/LSCNS/actions/runs/38112108205 |

Phase 4–8 통합 검증은 1440×900 데스크톱 Headless Chromium과 390×844 모바일 에뮬레이션에서 실행했습니다. 실제 클릭으로 10초 FPS 계산, MEP 경로 이격, NVIDIA 제품 카탈로그 JSON 로드, CPO 모듈 중복 제외, Ethernet+IB CPO 차단, Phase 8 내부 자가진단, 레거시 랙/장애 조작을 확인했습니다.

통합 실행의 10초 FPS 결과 6.97은 **GitHub Actions SwiftShader 소프트웨어 WebGL**에서 측정된 값입니다. 실제 Windows GPU, Android, iPhone의 성능 수치가 아닙니다.

## 4. 성능 변동성과 남은 과제

같은 CI 실행기에서 v4.6.1→v4.7.0 샘플 비교( https://github.com/kimheeseo/LSCNS/actions/runs/38111993953 ): Medium 25.2→26.3 FPS, Low 33.7→34.9 FPS. 공유 러너 자원 상황에 따른 변동이 커 실제 속도 개선이 통계적으로 확정됐거나 30/60FPS를 달성했다고 주장할 수 없습니다.

추가 필요: ① 실기기 3종 FPS JSON 기록 ② 진정한 GPU Instancing 셰이더/드로우 도입 검토 ③ 실측 MEP/BIM 간섭 ④ NVIDIA 주문 SKU·광모듈·MPO polarity·FEC·광예산 호환성 확인 ⑤ 기존 BOM 엔진과 정확한 수량 상호 동기화.

## 5. 사용 안내

[Phase 1~8 종합 가이드](../LS_Datacenter_3D_Phase8_Guide.md)에서 조작·입력·출력·검증 절차를 확인하십시오.

**결론:** v4.7.0은 Phase 4~8 연구용 기능의 실행 가능한 통합 구현입니다. 정식 COMSOL/BIM 설계 솔버, 실기기 성능 검증서 또는 구매 승인 BOM으로 표시하면 안 됩니다.
