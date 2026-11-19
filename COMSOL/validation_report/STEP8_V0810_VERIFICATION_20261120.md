# STEP8 v0.8.10 검증 / 한계 — 2026-11-20

STEP8은 소프트웨어 구조·실행·UI 개선이다. **STEP7 전체 물리 검증은 미완료**이며 V1/V4/V5/V8/V9 PASS, V2/V3/V6 근거 필요, V7 FAIL을 유지한다.

| 요구 | 반영 / 검사 | 범위·한계 |
|---|---|---|
| 계층 분리 | lswave/geometry, mesh, materials, solver, postprocess, benchmarks, reports, web | additive facade, legacy 파일/CLI/Colab 보존. 전체 앱 재작성 아님 |
| 공식/표시 분리 | formulas.json registry + Python common evaluator + separate ES display | signed loss/gamma/actual Gauss Aeff 공통 실행; legacy scalar JS는 유지 |
| 계산 근거 | actual 2모드 FEM의 α/Aeff, P/Q·Gauss 점 수·단위·출처 | 희소 field_samples는 정밀 적분하지 않음 |
| 메타데이터 | design+실행 옵션 manifest, SHA256, solver_version=0.8.6, UTC timestamp | canonical UTF-8 문자열 검증. 물리 인증/서명 아님 |
| 웹 표시 | 6성분 모드 장, actual labeled base n(x,y) map, STEP4 sweep/STEP5 gap 연결 | material map은 PM/bend tensor가 아닌 base scalar n. 기존 n(r) 뷰 유지 |
| 무효화 | 입력/새 결과 변경 시 basis·map·저장 결과 제거 | old results missing provenance는 경고 호환 |
| 재료 정합 | FSM/ARF benchmark air를 문헌/합성 조건 n=1로 명시 | 기존 air n=1.00027과 ambient unspecified provenance는 유지 |
| 로컬 시험 | namespace/metadata/tampering/unit/actual example tests | 전체 Python·Node 및 browser/CI 후속 증거 기록 |

[API·실행·한계](../universal_silica_v01/fiber_research_v1/STEP8_V0810.md) · [실제 예시](../universal_silica_v01/fiber_research_v1/benchmarks/step8_v0810/example_mode_result.json) · [STEP7 판정](STEP7_V089_VERIFICATION_20261120.md).

한계: V2 exact parameter file 및 V3 full-precision normalized bent CSV 미확인, V6 FSM/FM/HOM 수렴·모드 식별 부족, V7 공진 branch/peak 추적 실패. 새로운 자료 없이 PASS로 변경하지 않는다. 소프트웨어 구조 개선이 HCF 극소 손실·실제 PM 재료·정량 SMF-HCF splice loss 검증을 뜻하지 않는다. 다음 물리 작업은 V2/V3 자료 확보와 V6/V7 개선이다.
