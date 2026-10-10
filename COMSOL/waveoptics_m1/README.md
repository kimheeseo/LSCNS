# Python Wave Optics — M1

1D 대칭 슬랩의 TE/TM 모드, n_eff 및 필드 프로파일을 계산하는 실제 FEM 툴입니다.
Python 3.11 + NumPy/SciPy + gmsh + scikit-fem + matplotlib을 사용합니다.
M1만 구현했고 M2/M3/M4는 포함하지 않습니다.

## 설치

압축을 풀고 `waveoptics_m1` 폴더에서 실행합니다.

```bash
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install -e '.[test]'
```

Windows에서는 가상환경 활성화를 `.venv\Scripts\activate`로 바꾸고,
Python 3.11 선택에 `py -3.11`을 사용할 수 있습니다.
이 작업의 실제 검증 OS는 Linux이며 Windows 실행은 검증하지 않았습니다.

gmsh Linux wheel이 공유 라이브러리 오류를 내면 배포판 패키지로 설치합니다.
Ubuntu/Debian 예: `sudo apt-get install libglu1-mesa libxft2`.
그래픽 창은 열지 않습니다. gmsh 라이브러리를 사용할 수 없으면
`--mesh-backend native`로 같은 계면 정렬 1D 메시를 사용할 수 있습니다.
기본 수렴 보고서는 **gmsh 메시로 실제 실행**했습니다.

## 툴 실행

```bash
python -m waveoptics.cli --polarization both --out my_slab
```

생성 파일: `modes.png`, `modes.csv`, `fields_TE.csv`, `fields_TM.csv`, `run.json`.
콘솔에는 실제 계산한 n_eff, β 및 고유값 잔차를 출력합니다.
모든 길이 입력은 μm입니다. `width-um`은 코어 전체 두께, `padding-um`은 양쪽 각각의 클래딩 길이입니다.

```bash
python -m waveoptics.cli --n-core 3.45 --n-clad 1.44 --width-um 0.4 --wavelength-um 1.55 --padding-um 8 --h-um 0.0025 --out strong_slab
```

### Python API

```python
from waveoptics import solve_slab

s = solve_slab(n_core=1.5, n_clad=1.45, width_um=3.0,
               wavelength_um=1.55, padding_um=20.0,
               h_um=0.0375, polarization='TM', max_modes=6)
for m in s.modes:
    print(m.neff, m.beta_per_um, m.relative_residual)
# s.x_um: 메시 노드, m.field: 해당 노드의 L2 정규화된 H_y
```

TE 필드는 E_y, TM 필드는 H_y입니다. ∫|ψ|²dx=1로 정규화하므로 1 W 전력 정규화가 아닙니다.
β 단위는 μm⁻¹입니다. `s.mode_limit_reached`가 참이면 `max_modes`를 늘려 전체 모드 누락을 확인합니다.
새 조건에서는 `h_um`을 3단계 이상 줄이고 `padding_um`을 별도로 늘려 검증해야 합니다.

## 검증 재실행

```bash
python -m pytest tests -q --junitxml=results/pytest.xml
python validate.py
```

`validate.py`는 32개 테스트가 통과한 JUnit 결과를 요구합니다.
독립 해석해와 FEM을 다시 계산하여 표/CSV/JSON/PNG와 `M1_VALIDATION_REPORT.md`를 생성합니다.
기존 테스트 파일/초기 실패 로그/최종 결과 및 실행 버전은 함께 제공했습니다.
`results/environment_packages.txt`는 이번 실행의 기록이고, 설치 설정은 `pyproject.toml`입니다.

## 파일 구성

- `waveoptics/slab.py`: TE/TM 약형식 조립 및 일반화 고유값 풀이
- `waveoptics/mesh.py`: 실제 gmsh / native 계면 정렬 메시
- `waveoptics/cli.py`: 사용자 입력, 데이터 내보내기, matplotlib 그림
- `tests/analytic_slab.py`: 솔버와 독립인 무한 클래딩 해석해
- `tests/test_m1.py`: 솔버 구현 전에 작성한 28개 테스트
- `tests/test_additional_domain.py`: 경계 경고를 확인하는 추가 4개 테스트
- `validate.py`: 계산/수렴 표와 보고서 생성
- `M1_VALIDATION_REPORT.md`: 수식, 실제 수치, 검증된 범위 / 검증되지 않은 범위 / 알려진 한계
- `results/`: 실제 실행 로그, 원시 데이터와 그림

본 툴은 COMSOL 전체 대체품이나 실제 광섬유 2D 벡터 솔버로 검증된 것이 아닙니다.
복소 n_eff/손실, PML/포트/S-파라미터는 후속 마일스톤의 대상입니다.
