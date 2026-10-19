# M8: HCF 모세관 · 2코어 MCF 연구용 검증

이 도구는 **이상적 HCF 모세관의 스칼라 모드**와 **동일 원형 코어 2개의 스칼라 결합 모드**를 계산합니다. 실제 ARF/NANF 손실 또는 실제 MCF 랜덤 누화 검증까지 완료된 도구는 아닙니다.

[통합 결과 보고서](M8_VALIDATION_REPORT.md) · [HCF 보고서](HCF_VALIDATION_REPORT.md) · [2코어 MCF 보고서](MCF_2CORE_VALIDATION_REPORT.md)

[▶ Colab에서 실행](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/waveoptics_m8/M8_HCF_MCF_2Core_Validation.ipynb)

Colab에서 런타임 연결 후 **런타임 → 모두 실행**을 선택하십시오. `CONFIG`에서 입력을 바꾸면 재계산합니다. 엔진은 노트북에 내장되어 저장소 복제 없이 실행할 수 있습니다. Python 환경에서 내장 셀 재실행을 검증했으며, Colab 브라우저 UI의 수동 실행은 확인하지 않았습니다.

## 어떤 입력으로 무엇을 계산하는가

| 모델 | 주요 입력 (기본값) | 출력 |
|---|---|---|
| HCF 모세관 | λ=1.55 µm, 반경=15.0 µm, 일정 n_air=1.00027, 요소 수 | neff, 횡방향 kt², D [ps/(nm·km)], Aeff [µm²], 고유값 잔차·메시 오차 |
| 2코어 MCF | λ=1.55 µm, 반경=4.1 µm, Δn=0.005, 중심 간격=16.0 µm | 짝/홀 neff, Δneff, κ [1/m], Lc [mm], 슈퍼모드 장 |
| MCF 검증 | 목표 h=[1.0, 0.7, 0.45, 0.3] µm, 영역 반폭=[28.0, 32.0, 40.0] µm | 기준 대비 오차·메시/영역 수렴·간격별 비교 |
| MCF 전력 모델 | 전파 길이 z [m], Lc [mm] | 무손실 동일 코어의 P1/P2 전력 분율 |

예: 기본 2코어 조건에서 개선 FEM Lc=10.63981114 mm입니다. 전력 모델에서는 z=Lc에서 P2=1, z=2Lc에서 P2=0입니다. 이 수치는 실제 케이블 XT 스펙을 뜻하지 않습니다.

## 로컬 Python 실행

저장소 루트에서:

```bash
# Ubuntu/Colab의 Gmsh 공유 라이브러리
sudo apt-get install -y libglu1-mesa libxft2 libxrender1 libxinerama1 libxcursor1
python -m pip install -r COMSOL/waveoptics_m8/requirements-tested.txt
python COMSOL/waveoptics_m8/validation/run_validation.py
python -m pytest COMSOL/waveoptics_m8/tests -q
python COMSOL/waveoptics_m8/validation/make_report.py
```

원시 결과는 `validation/results/`에 CSV·JSON·PNG로 저장됩니다. 판정 실패 시 CLI는 오류 종료하며, 실패 수치를 저장합니다. 노트북은 실패 항목도 화면에 표시합니다. 전체 실행 시간은 이 환경에서 약 51초이며 하드웨어와 메시 크기에 따라 달라집니다.

### 직접 호출 예시

```python
from pathlib import Path
import sys
base = Path('COMSOL')
sys.path[:0] = [str(base/'waveoptics_m7'/'src'), str(base/'waveoptics_m8'/'src')]
from fiber_modes import Fiber
from mcf_hcf import solve_two_core_conforming
from multipole_reference import solve_reference
from capillary_fem import solve_capillary
from conforming_mcf import coherent_transfer

fiber = Fiber(wavelength_um=1.55, radius_um=4.1, delta_n=0.005)
mcf = solve_two_core_conforming(16.0, mesh_size_um=0.3, halfwidth_um=40.0, fiber=fiber)
ref = solve_reference(fiber, pitch_um=16.0, order=8)
hcf = solve_capillary(wavelength_um=1.55, radius_um=15.0, elements=640)
transfer = coherent_transfer(length_m=0.01, coupling_length_mm=mcf['coupling_length_mm'])
print(mcf, ref, hcf, transfer)
```

`solve_two_core`는 과거 Cartesian P1 호환 API입니다. 정밀 비교에는 `solve_two_core_conforming`과 독립 기준을 사용하십시오. `python COMSOL/run_m7_m8_validation.py`는 과거 M7/M8 시연을 실행하며, 새 검증의 대체 명령이 아닙니다.

## 오차 해석과 한계

오차율은 `100×|계산값−기준값|/|기준값|`이며, 유효굴절률 절대 오차는 `|Δneff|`입니다. **실제 제품·COMSOL·실측 데이터 대비 오차율은 아직 산정하지 않았습니다.** 오차율이 작아도 모델의 생략된 물리가 검증되는 것은 아닙니다.


MCF 입력은 동일 스텝형 두 코어·V<2.405·비중첩 조건을 사용합니다. 전체 벡터 편광, 비대칭/트렌치 코어, 굽힘·비틀림·손실은 포함하지 않습니다. 실리카 Sellmeier에 일정 Δn를 더하는 재료 가정입니다.
HCF는 J0/Dirichlet 모세관 근사입니다. 유리벽·반공진·PML·구속 손실은 미구현입니다.

## 파일

- `src/conforming_mcf.py`: 곡선 P2 FEM 및 전력 전달.
- `src/multipole_reference.py`: 독립 무한영역 스칼라 경계 정합.
- `src/capillary_fem.py`: HCF 모세관 방사형 P1 FEM.
- `validation/run_validation.py`: 수렴·독립 비교·그래프·판정 생성.
- `validation/results/summary.json`: 실행 조건과 모든 수치.
- `validation/build_notebook.py`, `replay_notebook.py`: 노트북 생성·실제 셀 재실행.
- `SOURCE_MANIFEST.json`: 소스 및 결과 SHA256.

버전 기록: [CHANGELOG.md](CHANGELOG.md).
