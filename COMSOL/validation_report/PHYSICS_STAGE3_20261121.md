# 3단계 — V6 FSM/FM/HOM 식별·수렴 (2026-11-21)

구현·실행 완료. V6 = **NEEDS_EVIDENCE 유지**, ESM_certified=false.

## 실제 계산

- λ=1200/1800 nm, Lambda=2 um, d/Lambda=.6에서 실제 periodic vector Maxwell FSM beta-root를 PWE cutoff 2/3/4/5/6으로 계산했다.
- finite PCF는 동일 물리 probe에서 다중 target으로 두 편광 FM 후보를 찾고, 세 메시 수준(EPW 4/5/6, curvature 64/96/128, bulk .8/.6/.45 um)을 부분공간 overlap으로 연결했다.
- rings=3/domain=8 um에서 rings=4/domain=10 um로 바꾼 별도 probe를 보존했다. 이 결합 변화는 독립적인 ring/window 수렴 PASS가 아니다.
- 각 파장에서 FM neff보다 .015/.035 낮은 target을 추가 탐색했다. FM 부분공간에 대한 실제 field projection, core Et proxy, 복소 neff, 편광/대칭 진단을 함께 남겼다. projection이 작고 neff가 낮은 후보는 HOM_CANDIDATE_NOT_IDENTIFIED로 분류하며 고유한 HE/TE/TM 명칭을 자동 확정하지 않는다.

## 결과

| λ [nm] | nFSM cutoff=6 | cutoff 5→6 변화 | FM 마지막 centroid | FM 마지막 메시 변화 | VPCF 후보 |
|---:|---:|---:|---:|---:|---:|
| 1200 | 1.3811588436004756 | 3.43387194e-4 | 1.4192356168236358 | 3.01598848e-5 | 3.4195447509 |
| 1800 | 1.3466333319188692 | 4.86980069e-4 | 1.3944469842204210 | 5.57931676e-5 | 2.5274012673 |

기존 1800 nm의 FM 탐색 실패를 회복했다. 1200 nm의 추가 탐색에서는 FM과 field projection이 약 .004–.021로 구분되는 lower-neff HOM 후보 record 4개를 얻었다. 독립 모드 개수·차수·cutoff는 확정하지 않았다. 1800 nm에서 해당 탐색 범위에 확인된 HOM이 없다는 사실은 HOM 부재나 cutoff의 증명이 아니다.

FSM 및 FM 마지막 두 절대 neff 변화 <=1e-5는 계산 전 선언한 공학적 진단 기준이다. 두 파장 모두 이 수렴 기준을 충족하지 못했다. 마지막 VPCF는 미수렴 후보값이다. 1800 nm에서 V<pi라는 수치만으로 ESM PASS를 만들지 않는다.

## 검증 / 원 자료

- 수렴 판정 2개 및 FM-subspace/HOM-candidate 분류 1개, 총 **3개 테스트 PASS**. 작은 algebraic residual을 수렴으로 대체하지 않는다.
- [요약](physics_20261121/V6/summary.json) / [수렴 CSV](physics_20261121/V6/convergence.csv).
- [실제 후보·장 ZIP](physics_20261121/V6/candidates.zip) / [멤버별 SHA256 index](physics_20261121/V6/candidate_index.json).
- 복원: `python package_physics_evidence.py unpack ../../validation_report/physics_20261121/V6`.
- 재계산: `OPENBLAS_NUM_THREADS=1 OMP_NUM_THREADS=1 python run_physics_v6.py --out /tmp/V6_fresh`.

## 남은 입력·실험·판정

cutoff 추가 확대/독립 FSM 검증, 더 높은 실제 메시 분해능, independent window/ring/PML sweeps, FM/HOM의 장·대칭/고유 차수 검토 및 HOM cutoff bracket, primary reference curve digitization와 불확도, 더 넓은 lambda/Lambda 및 d/Lambda 경계 조사가 필요하다. n_glass를 nFSM으로 대체하지 않는다. 상용 silica의 분산이나 실제 제조사 섬유를 인증한 계산이 아니다.
