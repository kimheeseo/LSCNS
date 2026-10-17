# Gold Anatomy DataCenter Tour — 8단계 개편 보고서

기준일: 2026-10-17  
대상: `DCI/DataCenter/LS_Datacenter_Gold_Anatomy.html`

## 목표
기존 Gold-style 탐방 UX는 유지하면서, 단순한 서버실 그림을 데이터센터 전체 해부도 형태로 확장한다. Utility/Power/Cooling/White Space/Network/MMR/Operations/Safety를 한 화면에서 이해하고, 조사 결과를 BOM/제품 툴로 연결한다.

## 8단계
1. **전체 맵/Zone 재설계**
   - Utility/Cooling Yard, Gray Zone Electrical, White Space IT Hall, Operations/Safety 영역 재배치
   - Gold-style 캐릭터 이동/조사 유지
2. **전력·냉각 흐름**
   - Utility → Transformer → ATS/Switchgear → UPS → PDU/Busway → Rack
   - A/B Feed 표현
   - Cooling Tower → Chiller/Pump → CRAH/CDU → Rack 및 Return loop
3. **White Space 고도화**
   - Rack Row A/B/C
   - Hot/Cold Aisle, Containment, 상부 Busway, Fiber/Copper Tray
4. **Network/MMR 고도화**
   - Carrier/WAN → ODF/Patch → Core/Border → Spine A/B → Leaf → Rack
   - Server↔Leaf, Leaf↔Spine, Spine↔Core/MMR 링크 가이드
5. **운영·보안·소방**
   - NOC/DCIM, Security Gate, CCTV, Access Control, Fire Detection/Suppression, Maintenance/Spares
6. **탐방 인터랙션/BOM 연결**
   - Junior OJT / Senior Review
   - Zone 빠른 이동
   - 조사 시 관련 BOM 및 DataCenter 제품 툴 링크
7. **모바일/컴퓨터 모드**
   - Auto / Mobile / Desktop
   - 모바일: 맵 중심 스크롤, 하단 제어, 선택 설비로 자동 이동
   - 데스크톱: 전체 해부도 + sticky 상세 패널
8. **검증 및 정리**
   - JavaScript 문법 검사
   - 5개 View, 3개 실행 모드, 2개 학습 모드, 전력/냉각/네트워크/운영 오버레이, 빠른 이동, BOM 링크 존재 확인

## 자동 정적 검증 결과
- JavaScript syntax: PASS
- 전체 구조 / 전력·냉각 / IT Hall / Network-MMR / 운영·안전 View: PASS
- Auto / Mobile / Desktop: PASS
- Junior OJT / Senior Review: PASS
- A/B Power Feed + CHW Supply/Return: PASS
- White Space Containment / Busway / Fiber-Copper Tray: PASS
- Carrier / ODF / Core / Spine / Leaf: PASS
- NOC / Security / Fire / Maintenance: PASS
- Keyboard 이동 / Z·Enter 조사 / Fast Travel: PASS
- BOM/Product Tool link: PASS

## 참고
이 페이지는 관심 유도와 교육/설계 검토를 위한 Gold-style concept visualization이다. 실제 CAD/BIM, 전기 단선도, 배관 P&ID, TIA-942/Uptime/ASHRAE 적합성 검증을 대체하지 않는다.
