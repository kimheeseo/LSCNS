> **2026-10-20 M8 개선:** [HCF 오차 보고서](waveoptics_m8/HCF_VALIDATION_REPORT.md) · [2코어 MCF 검증](waveoptics_m8/MCF_2CORE_VALIDATION_REPORT.md) · [실행 가능한 Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/waveoptics_m8/M8_HCF_MCF_2Core_Validation.ipynb). 곡선 P2 FEM·독립 기준·메시/영역 수렴을 추가했습니다. HCF는 이상적 모세관 모델 검증입니다.

# Python Wave Optics — M1–M8

## 입력값·출력값·예시·Mesh 비교 — 통합 Colab 보고서

[▶ G.654.E / G.657.A1 입력·출력 설명서](fiber_simulation_guide/readme.md) · [Colab 실행](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/fiber_simulation_guide/Fiber_Input_Output_Examples.ipynb) · [실행 결과 ZIP](fiber_simulation_guide/Fiber_Input_Output_Results.zip)

직접 Δn 입력과 MFD 보정 모드를 구분하고, 두 광섬유의 **실제 계산 예시 6개**, 분석 파장 변경, 코어 반경 변경, Mesh 수·외부 계산영역의 영향을 설명합니다. 사용자 Mesh·파장 입력을 바꾼 실제 재실행도 확인했습니다. [G.654.E 입력/출력](g654e_validation/readme.md) · [G.657.A1 입력/출력](g657a1_validation/readme.md)


## G.657.A1 Google Colab 실행 보고서 (2026-10-19)

[▶ Colab에서 G.657.A1 FEM 실행](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/g657a1_validation/G657A1_Colab_Validation.ipynb) · [검증 보고서](g657a1_validation/G657A1_VALIDATION_REPORT.md) · [그래프·표 결과 ZIP](g657a1_validation/G657A1_Colab_Results.zip)

1310 nm MFD 9.2 µm를 참고한 가상 단면의 radial/2D FEM·분산·영역/메시 수렴을 재계산합니다. **수치 검사 11/12 통과, 기존 2D 격자의 n_eff 기준 1건 미달. 규격 수치 참고 4/4 범위 내이며 굽힘 손실·실제 제품·ITU 적합성을 검증한 것은 아닙니다.** 저장된 그림 8개와 실행 결과가 포함돼 있습니다.


## 2026-10-19 연구용 수치 검증 업데이트

[연구용 실행·검증 안내](research/README.md) · [메시/영역/분산 검증 결과](research/results/BENCHMARK_REPORT.md)

원형 LP01 FEM, 정확한 장 적분, 무한구간 Bessel 기준 적분, MCF 수렴 검사와 회귀 검사를 추가했습니다. 검증 범위와 미검증 범위는 위 안내를 확인하십시오.


## M4 3D 전체 구현 보고서와 첨부사진 검증 추가

[2026-10-11 결과 보고서 및 재현 소스](reports/M4_3D_2026-10-11/)를 추가했습니다. 3D Nédélec·PML·직사각형 포트·S 행렬 구현의 기존 104개 실제 시험과 사진 기준 추가 검사 9개를 구분해 수록했습니다. 기존 아래 `waveoptics_m4` 공동 재구현은 유지하며 전체 주파수 응답 구현은 새 보고서 폴더의 소스 ZIP으로 제공합니다. 공동 주파수 0.1% 기준은 추가 세분화 메시에서 충족했으며 기존 기본 메시를 통과로 바꾼 것은 아닙니다. 3D 구 Mie 산란·임의 구조 전체 스퓨리어스 배제·사람 또는 별도 모델의 독립 검토는 미검증입니다.

> **두 가지 사용 방식 지원:** ① Python 소스를 직접 불러와 FEM 연구 및 파라미터 수정 ② 설치 없이 웹에서 광모드·분산·MCF/HCF 그래프 확인.
>
> **범위 주의:** Python M7은 약유도(weak-guidance) 2D 스칼라 FEM입니다. 웹은 Python FEM을 실행하지 않으며, LP01 독립 해석식, M8 저장된 FEM 결과의 보간, 이상적 HCF capillary 근사식을 사용합니다. 실제 COMSOL full-vector 기능·상용 광섬유 스펙 검증과 동일하다고 주장하지 않습니다.

## ① 웹페이지에서 파라미터 입력하고 그래프 보기

### **[▶ Wave Optics Lab · 웹 시뮬레이터 실행](https://kimheeseo.github.io/LSCNS/COMSOL/)**

- **M7 · Step-index SMF / G.654.E 개념:** 파장 \`λ\`, 코어 반경 \`a\`, 코어/클래딩 굴절률차 \`Δn\` 입력 → LP01 Bessel 해석식 \`n_eff\`, second-moment MFD, \`Aeff\`, 모드 분포, 파장별 \`n_eff\` 그래프. 입력된 형상은 임의의 원형 step-index이며 특정 제조사 G.654.E의 실제 프로파일을 재현한 것이 아닙니다.
- **M8 · MCF:** 검증된 M8 2코어 FEM 참조점(12/16/20/24 µm)을 코어 간격 기준으로 로그 보간해 \`Δn_eff\`, 결합 길이 및 그래프 표시. **웹에서 새 FEM을 계산하지 않습니다.**
- **M8 · HCF:** 공기 코어 반경·파장·공기 굴절률 → *ideal hollow capillary* 근사 \`n_eff\`, \`D\` 및 개념 모드 그림. NANF/ARF 반공진·PML·confinement loss는 계산하지 않습니다.
- **결과 저장:** 메인 그래프 PNG 및 CSV 다운로드.
- **원본 웹 소스:** [COMSOL/index.html](index.html) · [web/app.js](web/app.js) · [web/style.css](web/style.css)

웹페이지가 보이지 않을 경우 [GitHub Pages 배포 상태](https://github.com/kimheeseo/LSCNS/actions)에서 최신 배포가 성공했는지 확인하세요. 단순 정적 HTML/JS 페이지이므로 계산 서버와 새 외부 JavaScript 라이브러리를 요구하지 않습니다.

## ② 내 컴퓨터에서 Python FEM 직접 실행

### 환경 설치 (Windows PowerShell, Python 3.11 권장)

먼저 [Python 3.11](https://www.python.org/downloads/)과 Git을 설치합니다. 아래 명령어는 **PowerShell**에서 실행합니다.

```powershell
git clone https://github.com/kimheeseo/LSCNS.git
cd LSCNS
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install numpy scipy matplotlib
```

기존 검증 조건으로 M7(광섬유 FEM)과 M8(MCF·HCF) 그래프를 재생성합니다.

```powershell
.\.venv\Scripts\python.exe COMSOL
un_m7_m8_validation.py
```

실행 결과는 아래 파일로 저장됩니다.

- `COMSOL/waveoptics_m7/screenshots/` → `m7_mode_and_convergence.png`, `m7_dispersion.png`, `m7_dispersion.csv`, `m7_summary.json`
- `COMSOL/waveoptics_m8/screenshots/` → `m8_mcf_supermodes.png`, `m8_coupling_and_hcf.png`, `m8_summary.json`
- `COMSOL/waveoptics_m7/M7_VALIDATION_REPORT.md`, `COMSOL/waveoptics_m8/M8_VALIDATION_REPORT.md`
- `COMSOL/M7_M8_EXECUTION_SUMMARY.json`, `COMSOL/M1_M8_FINAL_REPORT.md`

M4–M6을 재실행하려면 다음 명령어를 사용합니다.

```powershell
.\.venv\Scripts\python.exe COMSOL
un_m4_m6_validation.py
```

macOS/Linux는 `python3 -m venv .venv`를 사용한 뒤 `.venv/bin/python`으로 실행할 수 있습니다.

### Python 코드를 직접 불러와 광섬유 사양 계산

저장소 최상위 폴더(`LSCNS/`)에 `my_fiber_test.py` 파일을 만들고 다음을 입력합니다.

```python
from pathlib import Path
import sys

# 연구용 Python 모듈 경로 추가
sys.path.insert(0, str(Path("COMSOL/waveoptics_m7/src").resolve()))

from fiber_modes import Fiber, solve_modes, scalar_metrics, lp01_analytic

fiber = Fiber(
    wavelength_um=1.55,        # wavelength (µm)
    radius_um=4.1,             # core radius (µm)
    delta_n=0.005,             # n_core - n_cladding
    domain_halfwidth_um=18.0  # simulation half-width (µm)
)

# 실제 Python 2D scalar FEM 해석. n은 홀수이며 15 이상.
solution = solve_modes(fiber, n=81)
print("2D FEM:", scalar_metrics(solution))
print("LP01 independent analytical reference:", lp01_analytic(fiber))
```

PowerShell에서 실행합니다.

```powershell
.\.venv\Scripts\python.exe my_fiber_test.py
```

**주의:** `radius_um`은 코어 **반경**(지름 아님), `delta_n`은 **굴절률의 절대 차이**입니다. 제조사의 MFD·Aeff·D만으로 실제 코어 굴절률 분포를 유일하게 결정할 수 없습니다. 단순 step-index 해석을 실제 G.652.D/G.654.E 제품의 내부 구조와 동일시하지 마세요.

### M8 2코어 결합 / 중공코어 근사 직접 불러오기

같은 스크립트 또는 별도 파일에서 아래 코드를 사용할 수 있습니다.

```python
sys.path.insert(0, str(Path("COMSOL/waveoptics_m8/src").resolve()))
from mcf_hcf import solve_two_core, capillary_neff, capillary_dispersion

mcf = solve_two_core(separation_um=16.0, n=101)
print("Core neff splitting:", mcf["index_split"])
print("MCF coupling length (mm):", mcf["coupling_length_mm"])

print("Ideal capillary neff:", capillary_neff(1.55, radius_um=15))
print("Ideal capillary D:", capillary_dispersion(1.55))
```

M8 MCF는 스칼라 FEM, HCF는 이상적 모세관 해석식입니다. **HCF의 complex `n_eff`, 누설 손실, 반공진 NANF 구조**는 현재 제공하지 않습니다.

## M1–M8 마일스톤 자료

| 단계 | 해석 목적 | 소스 · 검증 자료 |
|---|---|---|
| M1 | Slab scalar waveguide | [waveoptics_m1](waveoptics_m1/) |
| M2 | Vector-mode baseline | [waveoptics_m2](waveoptics_m2/) |
| M3 | TE/TM 2D scattering · PML · S parameters | [M3](waveoptics_m3/) · [검증 보고서](waveoptics_m3/M3_VALIDATION_REPORT.md) |
| M4 | 3D Nédélec PEC cavity 부분 재구현 | [M4](waveoptics_m4/) · [검증 보고서](waveoptics_m4/M4_VALIDATION_REPORT.md) |
| M5 | Material dispersion · spectral S (1D) | [waveoptics_m5](waveoptics_m5/) |
| M6 | Coating inverse design (1D) | [waveoptics_m6](waveoptics_m6/) |
| **M7** | Step-index optical fiber LP01 scalar FEM, MFD, Aeff | [M7](waveoptics_m7/) · [검증 보고서](waveoptics_m7/M7_VALIDATION_REPORT.md) |
| **M8** | Two-core MCF scalar FEM + HCF capillary reference | [M8](waveoptics_m8/) · [검증 보고서](waveoptics_m8/M8_VALIDATION_REPORT.md) |

M1~M3은 원본 소스와 수치 검증을 제공합니다. M4는 `waveoptics_m4`의 PEC 공동 솔버와 `reports/M4_3D_2026-10-11`의 전체 3D PML·포트 구현 보관 ZIP을 구분합니다. M5·M6은 1D 독립 해석 오라클입니다. MATLAB/COMSOL 소프트웨어 자체의 검증서나 상용 광섬유 제품 사양서와의 정량 비교는 아닙니다.

## 자동 테스트

- [M7·M8 실제 Python FEM 수치 검증](https://github.com/kimheeseo/LSCNS/actions/workflows/waveoptics-m7-m8-validation.yml)
- [웹 해석식·화면 smoke test](https://github.com/kimheeseo/LSCNS/actions/workflows/waveoptics-web-smoke.yml)
- [M1–M8 통합 검증 현황](M1_M8_FINAL_REPORT.md)

웹과 Python 계산 엔진은 명확히 분리돼 있습니다. 새 제품 사양을 비교할 때는 **Python FEM의 메시 수렴·해석해 오차**를 확인한 결과를 연구용 근거로 사용하고, 웹 미리보기 결과를 FEM 검증값으로 보고하지 마십시오.
