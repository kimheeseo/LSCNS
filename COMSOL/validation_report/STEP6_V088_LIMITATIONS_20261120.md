# STEP6 v0.8.8 한계 및 다음 검증 제안

2026-11-20. 앱 확장 성공은 광섬유 성능 인증과 다르다.

| 단계 | 남은 물리 검증 / 필요한 자료 |
|---|---|
| 6a | 2D 등각 굽힘과 COMSOL 축대칭 비교는 정식화 차이를 포함한다. 실제 섬유 단면/굴절률과 반경별 측정 손실, 독립 PML·창·메시 수렴이 필요하다. 총손실 차이는 pure bend-loss 증명이 아니다. R_eff/R 보정은 해당 섬유/파장/응력 조건의 문헌 계수 없으면 사용하지 않는다. |
| 6b | 실제 PANDA/Bow-tie rod 조성별 E·ν·α·p11/p12·열 이력·dn/dT, 측정 beat length 필요. 현재 예시는 합성 상수. 선형 generalized plane strain, full thermal strain 포함, 점탄성 이완/3D warping 미포함. 후보 pair도 독립 mesh/COMSOL 검증 필요. |
| 6c | 반환한 모드 수는 모든 guided supermode 확인이 아니다. 여러 target, polarization subspace 및 core power 확인 필요. κ는 동일 대칭 2코어의 유일한 편광/envelope 후보에 한정한다. 일반 N코어·굽힘 detuning의 임의 β 차이를 κ로 해석하지 않는다. Twist는 helicoidal 또는 longitudinal coupled-mode 모델 추가 필요. |
| 6d | Et overlap·phase tilt·homogeneous gap ASM·scalar Fresnel 근사. high-contrast 접속은 full E/H mode matching 필요. SMF↔HCF 융착 collapse, taper, 거칠기/표면 산란, back-reflection 및 다중 반사 미포함. 실제 예시의 finite grid edge power를 보고 창·sampling/padding 수렴 전 결과를 설계 손실로 사용하지 않는다. |
| 6e | 유한 grid 최적이며 전역 최적 보장 없음. 목적값은 실제 FE 후보 모드 집합의 값이다. 특정 FM/HOM 목적은 parameter-continuous branch 추적을 별도 확인해야 한다. Bayes/GA는 선택 요구이므로 구현하지 않았다. |
| 6f | 기존 Maxwell PMD 분포는 시연용이다. uniform signed Δn의 group DGD를 statistical PMD 계수와 동일시하지 않는다. longitudinal correlation length, stress/twist distribution, coupling statistics→Jones propagation→측정 DGD 분포 검증이 다음 단계다. |
| PCF V2/V6 | V2 공식 parameter txt가 필요하다. V6 문헌 7개 geometry/cutoff fit은 확인했지만 actual FEM FM/HOM cutoff는 미검증이다. 참고식 수치를 FEM 통과로 표시하지 않는다. |

다음 제안: V2 원 입력과 V6 cladding FSM/HOM cutoff 검증 → 실측/출처가 있는 PM rod 재료 검증 → SMF↔HCF full E/H 접속 모형. 이 후속 구현은 사용자 승인 후 진행한다.
