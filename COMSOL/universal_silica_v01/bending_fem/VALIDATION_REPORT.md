# Velamuri 2018 — 물리 기반 굽힘손실 2D FEM/PML 고도화 결과

> **2026-11-03 · 실제 해석 코드 추가:** [`solver.py`](./solver.py) · [Colab 실행](./Velamuri_Bend_2D_PML.ipynb) · [실행 예제 JSON](./example_validation.json) · **최종 상태: 수치 수렴 미완료 / 제품 검증 미수행**

## 1. 논문에서 무엇을 반영했는가?

Velamuri et al., *Study of Bending Losses in Optical Fibers using COMSOL* (COMSOL Conference 2018). 원문 [공식 PDF](https://cn.comsol.com/paper/download/579211/velamuri_paper.pdf).

| 논문에서 사용한 방법 | 구현 여부 | 설명 |
|---|---|---|
| 곡률에 따른 등가 굴절률 변환 | **구현** | n(x,y) → n(x,y)(1+x/R) |
| 2D 모드 고유값 직접 계산 | **구현(스칼라 P1 FEM)** | 2D 삼각형 FEM, 실제 복소 모드 계산 |
| 복소 PML 외부 흡수층 | **구현(복소 좌표 스트레칭)** | PML 두께·강도 사용자 입력 |
| Im(n_eff) → dB/turn 환산 | **구현** | 직경 d=2R 사용 |
| 2D LP₀₁ 모드 강도 이미지 | **구현** | 실제 FEM 모드 고유벡터 기반 |
| PML 두께·메시 재계산 | **구현** | 매 조건 독립 eigenvalue solve |
| 논문의 Full-vector EWFD | **미구현** | 현재 스칼라 편광 비구분 |
| GEBT 실제 응력 텐서 | **미구현** | 단순 빔·Poisson 근사 옵션만 |
| 실물 클래딩 62.5 µm 및 코팅 | **미구현** | 사각형 계산창의 조기 PML로 대체 |
| 실제 BIF 제조사 굴절률 프로파일 | **미확보** | 실측 중앙 dip / trench 필요 |
| 논문 실험값에 대한 독립 재현 | **미검증** | 논문의 BIF 모드/물성이 현재 가상 Step-index와 다름 |

## 2. 검증 사례 — 입력을 바꾸며 실제 계산

**시연용 Step-index 입력(논문 실물 BIF가 아님):**

`λ=1550 nm, a=4.1 µm, Δn=0.005, 굽힘 직경 d=10 mm, FEM 창 반경 22 µm, PML=10.85 µm (7λ), h≈0.7 µm, 응력광학 OFF`

| 항목 | 직접 계산값 |
|---|---:|
| 굽힘 `Re(n_eff)` | 1.446327526 |
| 굽힘 `Im(n_eff)` | 6.1073495×10⁻⁵ |
| 변환한 굽힘손실 | **67.55619 dB/turn (잠정치)** |
| 직선의 `Im(n_eff)` | −4.52635×10⁻¹² (수치 0 근처) |
| FEM h=0.80 µm | 70.84181 dB/turn |
| FEM h=0.65 µm | 72.18885 dB/turn |
| FEM h=0.55 µm | 68.92921 dB/turn |
| **메시 0.65→0.55 차이** | **4.729% → 2% 기준 미충족** |
| PML=4λ=6.20 µm | 62.72457 dB/turn |
| PML=6λ=9.30 µm | 69.92529 dB/turn |
| PML=7λ=10.85 µm | 67.55619 dB/turn |
| **PML 6λ→7λ 차이** | **3.507% → 2% 기준 미충족** |

따라서 이 사례의 **손실값은 수치적 수렴성을 확인하지 못했습니다.** 그 값이 물리적으로 옳다는 증거는 없습니다. 2%는 프로그램의 임시 진단 기준이지 논문이 규정한 품질 판정 기준이 아닙니다.

아울러 논문 BIF 모델은 **1550 nm / 10 mm 굽힘 직경**에서 시뮬레이션 **0.0119 dB/turn**, 실험 평균 **약 0.012 dB/turn**을 보고합니다. 현재 예제는 그 제품의 실측 RIP/응력 구조가 아니므로 수치가 다르다고 해서 논문의 결과와 직접 검증하거나 오류를 판정할 수 없습니다.

## 3. 해석식과 주요 주의점

`n_eq(x,y)=n_s(x,y)(1+x/R)`와 약유도 스칼라 2D FEM/PML을 이용해 복소 고유값 `β=k0 n_eff`를 얻습니다.

이때 `E~exp(iβz)`로 정의하면:

```text
alpha_dB/m = (20/ln10) · (2π/λ_m) · Im(n_eff)
L_one_turn = π·d_m
loss_dB/turn = alpha_dB/m · L_one_turn
```

음수 `Im(n_eff)`는 굽힘손실을 무조건 양수로 만들지 않고 계산 불가로 반환합니다. 모드 선택은 *코어 인근 에너지 분율*에 의존하므로 큰 곡률 및 누설 모드 교차 시 모드 연속 추적이 추가로 필요합니다.

**계산 경계:** 사각형 FEM 창 |x|,|y|=22 µm 안쪽에 물리적 클래딩·코팅 경계를 모델링한 것이 아니라 그 바깥부터 PML을 사용했습니다. 실제 125 µm 클래딩 외경은 별도 재료/기하 구조입니다. 이 단축 창 문제 때문에 현재의 `dB/turn`를 제조사 실험값과 비교하는 것은 적합하지 않습니다.

## 4. 활용 방법

1. [Colab 실행](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/bending_fem/Velamuri_Bend_2D_PML.ipynb) — Python SciPy 2D FEM 재계산, 벤딩/직선 모드 강도 및 메시·PML 스윕.
2. Colab에서 `bend2d_pml_result.json` 생성.
3. [Universal Silica Fiber Lab](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/?v=20261103-complex-pml-1#bend2d-fem)의 **02++ 실제 2D FEM/PML 결과**에서 JSON 파일 불러오기.
4. **메시/PML 변화율과 검증 상태**를 반드시 확인. 실패하면 경고를 그대로 유지하고 결과를 제품 성능으로 사용하지 않기.

기존 **측정값 기반 굽힘손실 보간**, **복소 n_eff 후처리**, **감쇠**, **PMD**, **1D LP₀₁** 기능은 변경하지 않았습니다.

## 5. 다음 단계 — 논문 재현을 위한 필수 추가 구현

- 실제 125 µm 클래딩과 코팅 및 비원형성 등 현실 광섬유 단면.
- GEBT 기반 굽힘 응력 텐서와 굴절률 텐서의 변형(단순 Poisson 근사 제거).
- 2D **Full-vector Maxwell** 모드 계산, 편광과 누설 모드 추적.
- 코어 주변 적응형 삼각형 메시와 PML 강도/두께/위치/외부 영역의 동시 수렴 시험.
- 논문의 **정확한 BIF 굴절률 실측 프로파일** 또는 공개된 실험 원자료를 확보한 후 실험값과 독립 비교.

**최종 결론:** 굽힘손실의 물리 모드 계산을 시작할 수 있는 **2D 복소 PML 연구 플랫폼**은 구현했지만, 논문과 동일한 물리 모델 및 신뢰 가능한 제품 손실 예측은 아직 달성되지 않았습니다.
