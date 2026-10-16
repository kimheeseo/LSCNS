# 데이터센터 구성 및 설계 계산 백서 — v1.4

작성 기준: 2026-10-16 · 독자: 케이블·광연결 및 데이터센터 개발 실무자

## 파일 접근

- [Word 백서 v1.4 (143페이지)](https://docs.google.com/document/d/1qBQ6A0QZEsUkErt4c5QxAjJEWztrPWPC/edit?usp=drivesdk&ouid=104119821971369669159&rtpof=true&sd=true)
- [사례 B 교차 검산 재현 ZIP](https://drive.google.com/file/d/1FZm-UFONl590ounB0jjQ52RrXQRqBqwm/view?usp=drivesdk)
- [GitHub 원본 Phase 7-A 계산 엔진](../../phase7-bom-engine.js)
- [GitHub 재현 테스트](./validate_caseB.cjs)

**보관 위치 주의:** Word·ZIP 바이너리는 Google Drive에 별도로 보관되어 있으며, **이 GitHub 저장소에 바이너리 파일이 업로드된 것은 아닙니다.** Drive 파일은 현재 소유자 전용으로 보관되어 있어 외부인은 접근 권한이 필요할 수 있습니다. GitHub에는 재현 테스트와 설명서를 기록합니다.

## v1.4 변경 내용

- 800G/1.6T, LPO/CPO, 800ZR/1600ZR 등 차세대 광연결 설계 보강
- 사례 B의 DGX H100 3경로 2+2+2 및 장애 시 성능 제한 대안
- 전력·냉각·광링크 수치 예제 및 민감도 분석 확대
- §94 공통 BOM 엔진 소스 교차 검산: 입력값, 실행 출력, 백서 계산값의 경계 구분
- 기존 100개 절 유지 및 목차 재정렬

## 사례 B 실행 재현

GitHub 저장소 루트에서 Node.js 18+ 설치 후:

```bash
node DCI/DataCenter/whitepaper/v1.4/validate_caseB.cjs
```

기대 출력: `PASS: 19 value assertions + 1 invalid schema rejection`

| 변수 | 기준값 | 구분 |
|---|---:|---|
| 랙 수 | 2 | 계산 도출이 아닌 입력값 |
| DGX H100 서버 | 8 | 명시된 입력값 |
| GPU | 64 | 8서버 × 8GPU/서버 |
| NIC 요약 | 8 | 공통 엔진의 기본 1개/서버 가정. 400G 실제 광포트 개수와 혼동 금지 |
| 활성 논리 광링크 | 64 | 예제 JSON 생성 가정 |
| 활성 링크 길이 합 | 1,920 m | 64링크 × 30m |
| 단일 링크 단선 후 | 63링크 / 1,890m | down 링크 1개 제외 |
| 랙 최대 IT 전력 | 41.8kW | 4서버 × 10.2kW + 부대 장비 1kW; 엔진 직접 출력 아님 |

**검증 범위:** GitHub의 `phase7-bom-engine.js` 수량 요약과 독립 전력 산식을 검산했습니다. 실제 GitHub Pages 화면 조작 결과, 광모듈 제조사 호환성, PSU 장애 시 운전성능, 냉각 실측 결과 등은 이 테스트로 보증하지 않습니다.

## 원본 무결성(SHA-256)

- Word: `8161dc006bc3b8aa2d29c9904e5ba73ea164b73a96b448baf7de05fc788a3628`
- 재현 ZIP: `e53428024c7cc0704e97f0c7270bcf83266e11ca7abf155375e1c00d873d1397`

백서는 교육·설계 검토용 자료로, 제조사 승인·발주·시공 검증을 대신하지 않습니다.
