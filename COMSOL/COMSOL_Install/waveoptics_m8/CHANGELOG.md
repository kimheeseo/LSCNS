# M8 변경 기록

## 2026-10-20 — HCF 수치 검증 · 2코어 MCF 개선

- HCF: 독립 방사형 P1 고유값 솔버 추가. Bessel 해 대비 neff·kt²·D·Aeff, 5점 파장 미분, 9개 반경/파장 조합 검사.
- MCF: 코어 경계를 따르는 곡선 P2 요소, 1/4 영역 짝/홀 대칭 조건 추가.
- MCF: 무한 클래딩의 독립 Fourier–Bessel 경계 정합 기준 추가. 전개 차수, 각도 샘플 수, 원거리 단일 코어 극한 검사.
- 메시·영역·코어 간격 수렴, 기존 Cartesian P1 비교, 실제 FEM 슈퍼모드 장 및 전력 전달 그래프 추가.
- 기존 `solve_two_core` 호환 API 유지. 사용자 Fiber를 지정할 수 있도록 개선. 권장 API는 `solve_two_core_conforming`.
- HCF/MCF 개별 보고서, 실행 결과가 포함된 독립 Colab, 입력 변경 재실행 검사, CSV/JSON/PNG/ZIP, 소스 해시와 읽기 전용 CI 추가.
- 실물 ARF/NANF 손실 및 랜덤 케이블 누화는 검증 범위에 포함하지 않음.

과거 `screenshots/`는 과거 시연 결과로 보존. 새 결과는 `validation/results/`에 저장.
