# 20260908 HCF2 loss validation

Petrovich HCF2 실측값을 fitting에 사용하지 않은 동심원 scalar leaky-mode 모델의 실행 검증입니다. 모든 예측은 기준값을 불러오기 전에 계산되며 `target_used_for_fitting=False`가 결과표에 기록됩니다.

## 결론

목표한 10–15% 독립 검증 오차는 달성하지 못했습니다. 7층 모델은 5층보다 크게 개선됐지만 Petrovich HCF2의 실제 비축대칭 double-nested 구조와 총손실을 대체하지 못합니다.

| 파장 | 단계 | 코드 leakage (dB/km) | 논문 실측 총손실 (dB/km) | signed error | absolute error | 10%/15% 통과 |
|---:|---|---:|---:|---:|---:|---|
| 1310 nm | Raw 5-layer | 210.408680 | 0.128 | +164,281.8% | 164,281.8% | 실패/실패 |
| 1550 nm | Raw 5-layer | 123.010202 | 0.091 | +135,076.0% | 135,076.0% | 실패/실패 |
| 1310 nm | Current no-fit 7-layer | 1.764759 | 0.128 | +1,278.7% | 1,278.7% | 실패/실패 |
| 1550 nm | Current no-fit 7-layer | 0.669577 | 0.091 | +635.8% | 635.8% | 실패/실패 |

- Raw 5-layer MAPE: **149,678.914%**
- Current no-fit 7-layer MAPE: **957.258%**
- 3-layer TMM 대 Bache TE 수학적 검산 MAPE: **3.010%**
- 공개 지름 범위 54개 계산 중 1개 root가 비수렴했으며 결과 CSV에 그대로 표시했습니다.

3층 검산의 3.010%는 같은 scalar/TE 문제의 구현 검증일 뿐 실제 DNANF 손실 정확도가 아닙니다.

## Raw / calibrated / hold-out 분리

| 구분 | 상태 | 데이터 사용 |
|---|---|---|
| Raw | 기존 5층 동심원 scalar TMM | Petrovich 목표값 미사용 |
| Current no-fit | 3개 유리막을 표현한 7층 동심원 scalar TMM | Petrovich 목표값 미사용 |
| Current calibrated | **생성하지 않음** | 독립적인 FEM/실측 training set이 없어 Petrovich 두 점 fitting은 금지 |
| Independent hold-out | 1310 및 1550 nm benchmark | 예측 완료 후에만 Petrovich 실측값과 비교 |

따라서 이 실행에는 calibration 학습 파장이 없고, 두 기준 파장 모두 target-excluded hold-out입니다. 향후 보정 모델은 다른 섬유/파장의 FEM training set으로 계수를 정한 뒤 Petrovich HCF2를 완전 hold-out으로 유지해야 합니다.

## 모델 수정과 참고 논문

1. M. Bache et al., **“Poor-man's model of hollow-core anti-resonant fibers”** (2019): 단일막 TE/TM/hybrid bouncing-ray 식과 bounce rate를 구현했습니다. 논문의 설계별 FEM 보정계수 `f_FEM`은 사용하지 않았습니다.
2. D. Bird, **“Attenuation of model hollow-core, anti-resonant fibres”** (2017): 공기/유리 동심 다층을 경계 연속조건으로 연결하는 모델을 반영해 3·5·7층 outgoing-wave 복소근 문제를 구성했습니다.
3. M. Petrovich et al., **“First broadband optical fibre with an attenuation lower than 0.1 decibel per kilometre”** (2025): HCF2 실측 1310 nm 0.128 dB/km, 1550 nm 0.091 dB/km를 benchmark로만 사용했습니다.
4. L. R. Murphy and D. Bird, **“Azimuthal confinement: the missing ingredient in understanding confinement loss in antiresonant, hollow-core fibers”** (2023): 동심원 모델이 빠뜨리는 방위각 구속을 실패 원인으로 반영했습니다.

## 입력값과 가정

- 코어 반경: 14.75 µm (기존 비교조건, 논문 공개 코어 지름 범위 안)
- 실리카 막 두께: 0.50 µm
- 큰/중간/작은 튜브 지름: 공개 범위의 중간값 31.05/23.75/7.70 µm
- 1차원 환산 공극: 6.30/15.05 µm
- 튜브 수: 5
- 실리카 굴절률: Sellmeier 식

공극은 측정값이 아니라 내부 접선형 동심원 surrogate 가정입니다. 공개 지름 범위의 최소/중간/최대 27개 조합도 target 기반 선택 없이 전부 계산했습니다.

## 10%를 넘기 때문에 필요한 개선

1. 실제 5-tube double-nested SEM contour를 사용하는 full-vector 2D complex-eigenvalue FEM과 mesh/PML 수렴성 검사를 구현합니다.
2. radial glass web과 Murphy–Bird의 azimuthal confinement를 포함합니다.
3. leakage-only 대신 `LL + surface-scattering loss + microbend loss + gas absorption` 총손실 모델을 구성합니다.
4. 실제 막 두께, gap, 타원도, 튜브 각도, 길이방향 변동, 표면 roughness PSD, 코팅·외경, 굽힘 조건을 확보합니다.
5. 보정이 필요하면 별도 FEM/실측 training fibre와 hold-out fibre를 분리하고 calibrated 결과로 명시합니다.

## 실행 파일

- `20260908_HCF_loss_validation_executed.ipynb`: 출력 셀과 그래프가 저장된 Colab
- `20260908_HCF_loss_validation_executed.html`: 실행 화면 HTML
- `run_20260908_validation.py`: 전체 재현 스크립트
- `hcf_concentric_tmm_20260908.py`: 수정된 solver와 Bache/Bird 계열 무보정 모델
- `results/`: CSV, JSON, Markdown, PNG/SVG 결과

수치 스크립트는 Python 3.12.13, NumPy 2.3.5, SciPy 1.17.0에서 직접 실행됐습니다. 실행 worker가 Jupyter 커널의 ZeroMQ 소켓을 차단하므로, notebook 출력은 동일 스크립트의 성공한 direct-CPython 실행 결과를 그대로 삽입했습니다. 노트북 셀은 Colab에서 다시 실행할 수 있습니다.
