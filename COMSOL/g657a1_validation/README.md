# G.657.A1 — 실행형 Colab FEM 검증 보고서

[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/g657a1_validation/G657A1_Colab_Validation.ipynb)

G.654.E와 동일하게 **보정 → Bessel 해석해 → radial FEM → 실제 2D FEM → 분산/메시/영역 비교**를 수행합니다. 저장된 실행 결과와 그림 8개를 노트북에서 바로 볼 수 있고, `런타임 → 모두 실행`으로 다시 계산할 수 있습니다.

- [Colab 보고서 파일](G657A1_Colab_Validation.ipynb)
- [검증 보고서](G657A1_VALIDATION_REPORT.md)
- [그림·CSV·JSON 결과 ZIP](G657A1_Colab_Results.zip)
- [소스·노트북·결과 전체 패키지](G657A1_validation_package.zip)

**기본 결과: 수치 검사 11/12 통과, 규격 수치 참고 4/4 범위 내.** 원형 FEM은 수렴하지만 기존 161×161 Cartesian FEM의 neff 오차 약 1.43e-5가 1e-5 기준을 넘습니다. 이 실패는 숨기지 않았습니다.

## 모델과 해석 범위

LS 2025 브로셔의 PremierLink A1/A2 MFD **9.2 µm @1310 nm**를 보정 목표로 참고합니다. 코어 반경 **4.1 µm는 가정**, Δn은 MFD에 맞춘 보정값입니다. 실제 A1 제품의 코어/trench 프로파일을 복원한 모델이 아닙니다. SimpleLink A1의 내부 구조나 사양을 이 값으로 단정하지 않습니다.

MFD는 Petermann-II와 second-moment를 구분하며, Aeff는 장 적분으로 직접 계산합니다. `n2=2.6e-20 m²/W`는 gamma 계산을 위한 사용자 변경 가능한 가정입니다. 분석 파장을 1550 nm로 바꿔도 MFD 보정 기준은 1310 nm로 유지합니다.

G.657.A1의 핵심인 **굽힘 손실, 감쇠, 케이블 컷오프, PMD**는 이번 직선·실수 굴절률 모드 해석기로 검증하지 않았습니다. 이상적 LP11 컷오프 proxy와 케이블 컷오프를 동일시하지 마십시오. 규격 범위 참고 비교는 ITU 적합성 인증이나 독립 제품 검증이 아닙니다.

## 로컬 재현

Python 3.12에서 시험했습니다. Colab은 대부분의 계산 라이브러리를 기본 제공합니다.

```bash
python -m pip install -r requirements-tested.txt
python g657a1_validation.py
python validate_notebook.py
```

일반 환경은 Jupyter 커널로 실행합니다. 로컬 커널 소켓을 사용할 수 없는 환경에서는 아래 명령으로 실제 IPython 셀 실행을 할 수 있습니다.

```bash
python g657a1_validation.py --in-process
```

이번 저장된 출력은 IPython에서 10개 코드 셀을 모두 실행해 생성했습니다. Colab 웹 UI의 수동 클릭 검증은 하지 않았습니다. [실행 방법·파라미터 검증 기록](results/notebook_validation.json)을 확인하십시오.

새 실행은 별도 결과 폴더를 사용하므로, 2D/민감도 계산을 끄면 과거 그림이 현재 결과 ZIP에 섞이지 않습니다. 분석 파장 1550 nm로 변경했을 때 Δn·1310 nm MFD 보정이 유지되고, 생략한 그림이 ZIP에서 제외되는 것도 실제 재실행으로 확인했습니다.

노트북을 수정했다면 `build_notebook.py`가 엔진과 작성 셀로 **새 노트북을 다시 생성하고 저장된 출력을 지운다**는 점에 유의하십시오. 다시 생성한 뒤 재실행하면 출력이 채워집니다.
