# 2단계 — V7 후보 진단 저장·branch 추적 (2026-11-21)

구현·실행 완료. 물리 판정은 **FAIL 유지**. 내부 공진 peak와 절대 loss는 인증되지 않았다.

## 변경

- 원본 7파장(981.75–1328.25 nm)과 실패 구간의 최대 1회 중간점 탐색을 실제 mixed-vector FEM으로 실행했다.
- 각 파장에 capillary seed 및 ±0.003, 이전 통과 pair의 평균 neff를 사용한다. screening을 통과한 pair가 나오면 추가 target을 생략하며, 실제 실행한 target만 기록한다.
- neff 근접값만으로 중복을 제거하지 않고 실제 E 장 상관으로 collinear 후보만 중복 처리한다. 두 독립 편광, 공간 overlap, core Et proxy 및 같은 부분공간의 correlation을 함께 검사한다.
- 모든 반환 후보의 복소 neff, signed loss, 잔차, weak Gauss, PML 비율, 분류, tracking E 장, probe 좌표/가중치, 탈락 이유, 입력 hash 및 mesh/solver 진단을 보존한다.
- 동일 입력·소스·probe의 계산 재사용은 출처와 SHA256를 표시한다. 후속 cache는 내용 hash에 따른 immutable snapshot을 보존해 원 출력의 덮어쓰기와 분리했다. 이 개선은 새로운 물리 계산을 의미하지 않는다.

## 결과

| λ [nm] | 판정 |
|---:|---|
| 981.75 | TRACKED_CANDIDATE |
| 1039.50 | FAILED |
| 1097.25 | FAILED |
| 1155.00 | FAILED |
| 1212.75 | FAILED |
| 1270.50 | FAILED |
| 1328.25 | FAILED |

원본 7점 외에 실패 시도/중간점 record 12개를 기록했다. 원시 후보 record는 294개로 **반복·재사용된 후보를 포함**하며 294개 독립 모드를 의미하지 않는다. 유효한 내부 peak는 null이고 ARROW 위치와의 정량 오차도 null이다. 첫 파장 외 구간에서 continuous core-polarization subspace를 확보하지 못했다.

1039.5 nm의 한 탐색에서는 잔차가 약 1e-14임에도 core Et intensity 비율이 약 4e-4–6e-4에 불과한 후보가 반환됐다. 작은 algebraic residual만으로 core FM이라고 판정할 수 없다. 다른 target·중간점에서 core-like 후보가 나타나도 고정된 overlap 기준을 통과하지 못했다. threshold를 낮춰 PASS로 바꾸지 않았다.

## 검증·산출물

- 모드 순서 교체, 잘못된 공간 branch, 중복 편광, branch 결측, 끝점/음수/비정상 spectrum, cache 입력·probe·snapshot, 무손실 압축/변조 검사를 포함한 신규 테스트 **6개 PASS**.
- [요약](physics_20261121/V7_adaptive/summary.json) / [원본 파장 CSV](physics_20261121/V7_adaptive/spectrum.csv).
- [원시 후보·실제 장 ZIP](physics_20261121/V7_adaptive/candidates.zip) / [멤버별 hash index](physics_20261121/V7_adaptive/candidate_index.json).
- ZIP은 숫자를 반올림하지 않은 JSON과 재시작 시 acquisition source snapshot을 보존한다. acquisition의 일부는 cache 개선 전 코드로 계산됐으므로 개별 JSON의 source hash가 현재 helper와 다를 수 있다. solver/geometry/material 소스는 동일하며 계산을 새로 했다고 표시하지 않는다.

복원: fiber_research_v1 디렉터리에서 `python package_physics_evidence.py unpack ../../validation_report/physics_20261121/V7_adaptive`.
새 계산: `OPENBLAS_NUM_THREADS=1 OMP_NUM_THREADS=1 python run_physics_v7.py --out /tmp/V7_fresh --max-bisections 1`.

## 남은 한계

실수 target neff 기반 탐색으로 공진 주변의 동일 core FM branch가 확보되지 않았다. 더 넓은/복소 shift 지원, 혼성화 구간의 장 식별 및 적응형 분해능, 독립 mesh/PML/window 검증이 필요하다. 현재 Et-disk intensity proxy와 0.8 correlation 정책은 FM 정체성·전력 비율·수렴의 독립 증명이 아니다. V7 FAIL은 ARROW 이론의 반증이 아니다.
