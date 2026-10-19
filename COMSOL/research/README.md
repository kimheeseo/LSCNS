# Wave Optics research qualification · 2026-10-19

이 폴더는 기존 M1–M8 보고서를 보존하면서 **현재 소스를 다시 실행한 연구용 검증 근거**를 관리합니다. 지원한 모델과 검증한 조건 안에서만 연구 결과로 사용하십시오.

## 이번 개선

- M7: P1 삼각형 내부 7점 적분으로 유효면적·중심 기준 second-moment MFD 계산. 기존 노드 lumping에 따른 적분 오차 제거.
- M7: 원형 코어 경계에 메시를 맞춘 `radial_fem.py` 추가. 2D Cartesian FEM과 다른 이산화로 LP01 해석해에 대조. 원점 자연경계·외곽 Dirichlet.
- M7: Bessel 기준장의 적분을 코어/클래딩 분리 및 무한구간 적응 적분으로 변경. 기준값이 FEM 영역 크기에 의존하던 문제 제거.
- M7: 고유모드 초기벡터 고정, 입력 유한성·형상·모드 수 검사.
- M8: 이상적 capillary 분산의 수치 미분을 정확한 2차 도함수로 교체. 반경·공기 굴절률 인자를 지원.
- M6: `reflectance()` 기본 두께 계수 `0`을 `1`로 수정. 이전 최적화는 명시적인 0.97/1/1.03 계수를 사용했으므로 그 결과는 이 오류의 영향을 받지 않았습니다.

## 재현

저장소 루트에서 실행합니다. Python 3.12 가상환경을 권장합니다.

```bash
python -m venv .venv
# Linux/macOS: source .venv/bin/activate
# Windows PowerShell: .\.venv\Scripts\Activate.ps1
python -m pip install -r COMSOL/research/requirements.txt
python -m pytest COMSOL/research/tests -q
python COMSOL/research/run_benchmarks.py
python COMSOL/research/run_regressions.py
# 선택: 보관된 M4 전체 3D 핵심/추가 시험까지 재실행
python COMSOL/research/run_regressions.py --include-archived-m4
```

Linux에서 M1–M4를 실행할 때 gmsh의 OS 라이브러리가 필요합니다: `sudo apt-get install libglu1-mesa libxft2 libxrender1 libxinerama1 libxcursor1`.
수치 스레드 재현성을 위해 실행 전에 `OPENBLAS_NUM_THREADS=1`, `OMP_NUM_THREADS=1`을 설정하십시오.

- [원시 수치 및 환경·소스 해시](results/benchmarks.json)
- [수렴·독립 해석해 대조 보고서](results/BENCHMARK_REPORT.md)
- [회귀 실행 기록](results/regression_summary.json) — 전체 검사가 끝난 후 생성

### 연구용 원형 광섬유 계산

```python
from pathlib import Path
import sys
sys.path.insert(0, str(Path('COMSOL/waveoptics_m7/src').resolve()))
from fiber_modes import Fiber
from radial_fem import solve_radial, radial_dispersion

f = Fiber(wavelength_um=1.55, radius_um=4.1, delta_n=0.005,
          domain_halfwidth_um=40.)
result = solve_radial(f, core_elements=320, clad_elements=1280)
print({k: result[k] for k in ('n_eff','Aeff_um2','MFD_2nd_um','residual')})
print(radial_dispersion(f, step_um=.005, core_elements=320, clad_elements=1280))
```

새로운 조건마다 메시·외부 경계·파장 간격의 수렴을 확인하십시오.

## 검증에서 확인한 주의점

Cartesian FEM은 원형 경계의 굴절률을 삼각형 중심에서 분류하기 때문에 메시를 세분화해도 n_eff 오차가 단조 감소하지 않습니다. 81격자 한 점에서 작은 오차가 나왔다고 전체 수렴이 입증되지는 않습니다. 원형 광섬유의 정밀 연구에는 경계 정합 radial FEM을 권장하며, Cartesian FEM의 잔차는 선형대수 해의 정확도만 보여줍니다.

회귀 및 벤치마크는 실패 시 비정상 종료합니다. GitHub Actions는 실패를 포함한 원시 증빙을 artifact로 보관하고, 실행 결과를 자동으로 main에 덮어쓰지 않습니다. 이전 M7/M8 runner가 전체 M1–M8 상태 보고서를 덮어쓰던 동작도 제거했습니다.

## 검증의 의미와 사용 제한

| 항목 | 연구에 사용할 수 있는 범위 | 추가 검증이 필요한 범위 |
|---|---|---|
| 원형 광섬유 | 보고서의 네 step-index 조건에서 LP01 n_eff·Aeff·second-moment MFD·분산 검증 | 실제 도핑 분포, trench, 제조사 제품 스펙, 온도·응력 |
| 2D Cartesian FEM | 기존 기본 형상 및 메시/영역 검사 결과 | 작은 잔차만으로 형상·물리 정확도를 보장할 수 없음 |
| MCF | 동일한 직선 2코어의 초모드 분리·이상적 전력전달 길이, 지정 pitch의 수치 변화 검사 | 실제 케이블 XT, 굽힘·비틀림·랜덤 결합·다중코어 통계 |
| HCF | 일정한 n_air를 가정한 이상적 모세관 해석식 | NANF/ARF 반공진, 복소 n_eff, confinement loss |
| M1–M6 | 재실행 로그에 명시된 기존 수치/해석 예제 | 임의 구조의 정확도, 전 범위 COMSOL 동등성 |

MFD_2nd는 중심 기준 스칼라 강도 2차 모멘트 직경입니다. 제조사 MFD 또는 Petermann-II 직경과 같다고 가정하지 마십시오. Aeff는 장분포에서 직접 적분하며 `pi*(MFD/2)^2`와 일반적으로 같지 않습니다. 모델은 n_core(λ)=n_silica(λ)+일정 Δn을 가정하므로 실제 Ge-doped silica 물질 분산을 재현했다고 볼 수 없습니다.

해석해 대조는 **수치 검증(verification)**입니다. 측정값을 통한 **실물 검증(validation)**과 구분하십시오. 실제 COMSOL 실행 결과 또는 광섬유 측정 데이터와의 직접 대조는 이번 작업에 포함하지 않았습니다.

## 공식 문헌 및 후속 대조 기준

- Malitson (1965), fused-silica Sellmeier: https://doi.org/10.1364/JOSA.55.001205
- COMSOL step-index fiber 공식 예제: https://doc.comsol.com/6.3/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html
- COMSOL microstructured leaky-mode 공식 검증 예제: https://doc.comsol.com/6.3/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html

공식 예제 링크는 모델의 범위와 향후 대조 절차를 위한 출처입니다. 이 링크의 결과를 이번 코드의 직접 대조 결과로 취급하지 않았습니다.
