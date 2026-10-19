# G.657.A1 입력·출력 및 실행 예시

[통합 Colab 실행형 보고서](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/fiber_simulation_guide/Fiber_Input_Output_Examples.ipynb) · [전체 입력/출력 설명서](../fiber_simulation_guide/readme.md)

## 기본 입력

- 분석 파장: 1310 nm
- MFD 기준 파장: 1310 nm
- 가정 코어 반경: 4.1 µm (지름 아님)
- 목표 Petermann-II MFD: 9.2 µm
- MFD 보정 모드: True. Δn을 목표 MFD에 맞춰 계산합니다.
- Radial 코어 요소 수: 40/80/160/320; 클래딩 요소 수는 4배; 외곽 반경 40 µm 고정.
- 2D 격자: 61/81/121/161; 계산영역 반폭 24 µm 고정.

## 이 입력의 실제 출력

| 출력 | 값 |
|---|---:|
| 보정 Δn | 0.0046215295 |
| n_eff, Bessel 기준 | 1.449107548135 |
| MFD, Petermann-II | 9.200000 µm |
| MFD, second-moment | 9.420343 µm |
| Aeff | 66.267872 µm² |
| D | 0.560432 ps/(nm·km) |
| gamma | 1.881824 1/(W·km), n2 가정 사용 |

MFD는 보정에 사용했으므로 독립 검증값이 아닙니다. 분석 파장을 바꿔도 MFD 기준 파장은 유지합니다. 직접 계산 모드(`fit_delta_n_from_mfd=False`)에서는 입력 Δn으로 MFD·Aeff·분산을 예측합니다.

## Mesh에 대한 주의

위 모드 특성은 이론해 기준 출력입니다. 실제 FEM 값은 Mesh 수와 계산영역에 따라 변하므로 Colab의 Mesh 비교 표에서 별도로 확인하십시오. 단일 Mesh의 오차율을 솔버 전체 정확도로 해석하지 마십시오.

실제 제품 굴절률 구조·감쇠·굽힘 손실·케이블 cutoff·PMD·COMSOL 동등성은 검증하지 않았습니다. 기존 제품별 검증 보고서는 [README.md](README.md)를 참조하십시오.
