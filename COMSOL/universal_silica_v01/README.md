# Universal Silica Fiber FEM Lab — v0.1 (Google Colab)

[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/Universal_Silica_Fiber_v0_1.ipynb)

**실리카 광섬유 단면 설계 → FEM 모드 해석 → 광특성·분산·메시 수렴성 보고서**의 첫 번째 버전입니다.

## 사용 방법
1. 위 **Open in Colab** 링크를 엽니다.
2. `런타임 → 모두 실행`으로 저장된 기본 예시를 재계산합니다.
3. `2. 입력 모델`에서 `profile`을 `step`, `graded`, `trench`, `multilayer` 중 선택하고 재료 파라미터를 수정합니다.
4. 결과표·모드 단면·파장별 D·FEM 메시 수렴을 검토합니다.
5. 출력 폴더의 `Universal_Silica_Fiber_Output.zip`에서 CSV·JSON·PNG를 내려받을 수 있습니다.

독립 실행 소스: [`universal_radial.py`](universal_radial.py). 기본 실행 결과: [`Universal_Silica_Fiber_Output.zip`](Universal_Silica_Fiber_Output.zip). 수치 검증 및 한계: [`VALIDATION_REPORT.md`](VALIDATION_REPORT.md).

## 지원 범위
| 구분 | v0.1 상태 |
|---|---|
| Silica Step-index 원형 단일코어 | **실제 radial P1 FEM** + 독립 LP01 Bessel oracle |
| Graded-index | **실제 radial P1 FEM**, α 파라미터 직접 입력 |
| Trench-assisted | **실제 radial P1 FEM**, 트렌치 위치·Δn 입력 |
| 다층 원형 실리카 | **실제 radial P1 FEM**, 반경·층 Δn 배열 입력 |
| 광학 출력 | neff, LP01 모드, Petermann-II 및 2차모멘트 MFD, Aeff, gamma(가정 n2), D(5-point finite difference) |
| 실제 비원형 2D 형상·MCF | **미지원**, 기존 별도 M8 2코어 솔버 존재 |
| NANF/ARF HCF · PCF | **미지원**, full-vector FEM + PML 필요 |
| 편광/복굴절·굽힘·누설손실 | **미지원** |

`n_clad(λ)`는 순수 실리카 Sellmeier, 모든 지역 Δn은 **파장 독립 상수**라는 가정입니다. 도핑 Ge/F별 고유 Sellmeier, 열·응력 의존성, 실제 제품 측정 굴절률 분포는 구현되지 않았으므로 제조사 G.652/G.654/G.657 적합성 검증에 사용하면 안 됩니다.

## 정확도 정책
- 계산영역 외벽은 가상 Dirichlet 경계이며 **실제 125 µm 외경/클래딩 경계가 아님**.
- `Step-index`이고 V<2.405이면 독립 Bessel LP01 해석해와 비교합니다.
- 나머지 단면 구조는 메시·영역·파장 미분 간격 테스트만 수행하며 독립 논문/COMSOL 데이터 기반 정확도 검증은 **별도 과제**입니다.
- 저장된 그래프는 수치해석 결과이며, 축대칭 2D 모드 그림은 계산된 **radial FEM 필드를 회전시킨 시각화**입니다. 2D full-vector FEM은 아닙니다.

## 다음 단계 계획
- v0.2: 실제 굴절률 측정 CSV 업로드, 분산 물성 데이터, LP0m/LP1m 고차 모드, 설계 파라미터 스윕.
- v0.3: Gmsh 기반 2D interface-conforming FEM, 2D 형상 도구, 2-core MCF.
- v0.4: Full-vector Maxwell + complex neff + PML + leakage / bending models.
- v0.5: NANF/ARF HCF, PCF, PMF, 제조사 독립 기준과 비교하는 벤치마크.

**최종 목표가 모든 실리카 기반 구조를 다루는 것이지, v0.1에서 이미 모두 지원한다는 의미는 아닙니다.**
