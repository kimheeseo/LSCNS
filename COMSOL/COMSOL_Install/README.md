# COMSOL_Install — 기존 Wave Optics 연구·설치 소스

그동안 COMSOL 루트에 흩어져 있던 **M1–M8 단계별 광학 해석 소스**, 기존 **웹 시뮬레이터(web)**, **Python FEM Workbench(workbench)**를 한곳에 정리했습니다.

| 폴더 | 내용 |
|---|---|
| `waveoptics_m1` ~ `waveoptics_m8` | 광학 모델 및 FEM 검증 단계별 코드 |
| `web` | 기존 Wave Optics Lab JavaScript · CSS · 테스트 |
| `workbench` | Python FEM Workbench 소스 · 배포 ZIP |

현재 범용 실리카 광섬유 도구는 [실행하기](../universal_silica_v01/web/)에서, 제품 검증 보고서는 [validation_report](../validation_report/)에서 확인할 수 있습니다.

[COMSOL 메인 안내](../README.md) · [COMSOL 웹 대시보드](../index.html)

## 이전 루트 파일 정리 · 2026-11-03

루트에 혼재하던 M1–M8 아카이브, 두 개의 Python_WaveOptics ZIP, JSON 요약,
[기존 대시보드](index.html), [광섬유 입력·출력 가이드](fiber_simulation_guide/),
기존 .gitignore를 이 폴더로 이동했습니다. 바이너리 ZIP/DOCX와 각 파일의 Git blob은 변경하지 않았습니다.

- [M1–M8 최종 보고서](M1_M8_FINAL_REPORT.md)
- [M2 광모드 분석 보고서 DOCX](M2_Optical_Mode_Analysis_Report.docx)
- [M7/M8 실행 결과 JSON](M7_M8_EXECUTION_SUMMARY.json)
- [M4–M6 실행 결과 JSON](WAVEOPTICS_M4_M6_EXECUTION_SUMMARY.json)
- [광섬유 입력·출력 검증 가이드 및 Colab](fiber_simulation_guide/README.md)

대표 문서는 상위 [COMSOL/README.md](../README.md) 하나로 유지합니다.
