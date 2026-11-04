> **2026-11-04 · COMSOL GUI형 화면 업그레이드:** [웹 실행](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/?v=20261104-comsol-style-ui-v1) · [구현 설명](universal_silica_v01/web/COMSOL_STYLE_README.md) · 기능 보존: G.654.E/COMSOL Step-index 검증, HCF, MCF, Full-vector Colab, PMD, 3D BPM. 실제 솔버 코드는 수정하지 않고 리본, Model Builder, Settings, Graphics, Add Study, Messages 패널과 밝은 테마/모바일 반응형 추가.

> **2026-11-04 COMSOL 공개 Step-index 검증:** [웹에서 직접 실행](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/?v=20261104-comsol-step-v2) · 기존 Scalar P1 FEM으로 n_eff 1.444401819, 공식 공개값 1.4444 대비 상대오차 0.000126%로 5% 미만. [검증 보고서](validation_report/comsol_step_index_benchmark/WEB_PRESET_VALIDATION.md). 공식 COMSOL 값은 4자리 공개치이며 제품 G.654.E 검증을 뜻하지 않음.

> **2026-11-04 설계 예제 라우팅 업그레이드:** [기존 툴](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/?v=20261104-preset-hcf-vector-v1) 왼쪽의 **설계 예제**에서 **HCF/NANF** 또는 **Full-vector(Colab)**를 선택해 전용 입력·결과 패널로 전환합니다. HCF는 실시간 튜브 단면/ARROW 근사 + 별도 Scalar complex FEM/PML 결과 import(미수렴), Full-vector는 **실리카용 2D FDFD Colab 결과 JSON**의 neff·MFD·Aeff·Ex/Ey/Ez 시각화입니다. **HCF Full-vector FEM은 미구현**. [사용 가이드](universal_silica_v01/web/README.md).

> **2026-11-04 HCF 복소 Scalar FEM + PML 실험 모듈:** [웹에서 FEM 결과 보기](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#hcf-fem) · [실제 FEM 계산 Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/hcf_complex_fem/HCF_Complex_FEM_Research.ipynb) · [검증 보고서](universal_silica_v01/hcf_complex_fem/VALIDATION_REPORT.md). 복소 neff 및 파장별 FEM 국소 extrema 후보를 계산합니다. **2D Scalar FEM 연구용이며 Full-vector 아님, 수렴검증 실패로 dB/km 미검증**.

> **2026-11-04 · 3D 광장 애니메이션 개선:** [3D 전파 뷰어](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#fiber-3d)에서 BPM 계산 후 발광 선단·단면 링 자동 재생, 속도 조절 및 3D 회전. [WebGL 소스](universal_silica_v01/web/fiber_3d_webgl.mjs) · [검증](universal_silica_v01/web/fiber_3d_webgl/VALIDATION_REPORT.md). 실제 계산은 2+1D Scalar BPM, 펄스 효과는 시각화용이며 3D Maxwell/HCF 손실 계산 아님.

> **2026-11-04 · 진짜 3D 광섬유 애니메이션 (WebGL2):** [3D 광 전파 실행](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#fiber-3d) · [WebGL 렌더러 소스](universal_silica_v01/web/fiber_3d_webgl.mjs) · [테스트·한계](universal_silica_v01/web/fiber_3d_webgl/VALIDATION_REPORT.md). 2+1D BPM 광장을 실제 x/y/z 3D 좌표로 렌더링·카메라 회전·확대·재생하며, HCF 3D 구조 시각화도 포함. 기존 2D로 전환 가능. **3D Full-vector Maxwell 해석이나 HCF 손실 계산은 미구현.** GPU 픽셀 검증은 실행 환경에서 WebGL이 차단되어 미확인.

> **2026-11-04 · 3D 광섬유 구조 및 빔 전파 업그레이드:** [실시간 3D 구조·BPM 웹 시뮬레이터](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#fiber-3d) · [3D 전파 Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/propagation_3d/Universal_Silica_Fiber_3D_Propagation.ipynb) · [수치 검증](universal_silica_v01/propagation_3d/VALIDATION_REPORT.md). Solid silica Step/Graded/Trench/Multilayer **2+1D Scalar FFT BPM 실제 z 전파**; HCF는 Nested Tube 3D 구조만 시각화. **3D Full-vector Maxwell FEM/FDTD 및 HCF 누설손실 미구현.**

> **2026-11-04 · Full-vector 모드 업그레이드:** [Colab Full-vector FDFD 실행](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/full_vector/LS_Full_Vector_Fiber_v0_1.ipynb) · [기존 웹에서 계산 결과 가져오기](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#full-vector-tool) · [독립 HE11 검증](universal_silica_v01/full_vector/VALIDATION_REPORT.md). 2D 결합 H 벡터 고유모드에서 neff, Ex/Ey/Ez 장 적분 MFD/Aeff 계산. **FDFD이며 FEM 아님.** 웹은 Colab 결과 스냅샷 뷰어; 직접 Full-vector 계산/복소 PML/누설손실 미지원.

> **2026-11-04 · HCF Nested Tube 설계 UI:** [**HCF Designer 실행**](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/#hcf-designer) · [코어·튜브 수·비율·벽 두께 동적 단면과 광특성](universal_silica_v01/web/hcf_designer/README.md) · [테스트 보고서](universal_silica_v01/web/hcf_designer/VALIDATION_REPORT.md). **공진식+이상적 모세관 근사**, 실제 중첩 구조의 벡터 FEM/손실 dB/km 직접 예측 아님.

> **2026-11-03 · MCF Cross-talk 웹 계산 추가:** [Universal Silica MCF XT 분석](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/) · [구현·가정 설명](universal_silica_v01/web/README.md). 2코어 M8 FEM 참조값의 로그 보간과 무손실 coherent 결합을 통한 XT(dB), P₁/P₂, κ, Lc 계산이며, **실제 제품 누화 실측 또는 XT(dB/km) 아님.**

> **폴더 정리 · 2026-11-03:** 메인 안내문은 이 `README.md` 하나로 통일했습니다. 기존 연구 산출물·ZIP·구형 웹페이지·Fiber I/O 가이드는 [COMSOL_Install](COMSOL_Install/)에 보관합니다. 기존 `/COMSOL/` 웹 주소는 설치 폴더의 대시보드로 자동 이동합니다.

> **2026-11-03 · 실시간 웹 UI v0.2:** [Universal Silica Fiber Workbench 실행](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/) · [소스·입력 가이드](universal_silica_v01/web/README.md) · [Python 기준 대조 결과](universal_silica_v01/web/VALIDATION_REPORT.md). GitHub Pages 브라우저에서 원형 Step·Graded·Trench·Multilayer 스칼라 FEM을 **실제로 다시 계산**합니다. 입력값을 바꾼 후 새 계산 실행 필요하며 결과 JSON/CSV는 로컬 다운로드입니다. **실제 LE125 인증/Full-vector/손실/임의 2D 구조 미지원.**

> **2026-10-30 · Universal Silica Fiber Colab v0.1:** [직접 단면을 설정하고 FEM 실행](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/Universal_Silica_Fiber_v0_1.ipynb) · [지원 범위/한계](universal_silica_v01/README.md) · [수치 검증](universal_silica_v01/VALIDATION_REPORT.md). Step-index·Graded·Trench·동심원 다층 **스칼라 radial LP01 FEM만 지원**, 비원형 MCF/HCF/PCF·full-vector·손실은 미구현.\n\n> **공개 자료 기반 비교 완료:** [M1·M7·MCF·HCF 결과·오차·한계](public_reference_benchmarks/results/COMPARISON_REPORT.md) · [재현 코드·출처](public_reference_benchmarks/README.md). 공개 COMSOL 예제의 조건과 논문을 재현했으며, COMSOL 프로그램 자체를 실행한 비교는 아닙니다.

> **설치·사용·결과 해석 가이드:** [이 도구로 알 수 있는 것, Windows 설치, 화면 사용 순서, 실제 계산 예시 3개](COMSOL_Install/workbench/USER_GUIDE.md) · [입력 예제 JSON](COMSOL_Install/workbench/examples/README.md)

> **2026-10-20 Workbench UI:** [모델 트리·설정·메시·결과 화면](https://kimheeseo.github.io/LSCNS/COMSOL/COMSOL_Install/workbench/) · [회사 PC에서 실제 Python FEM 실행하기](COMSOL_Install/workbench/README.md). GitHub 화면은 저장 결과 미리보기이며 새 해석은 로컬 Python 앱에서 실행합니다.

> **2026-10-20 M8 개선:** [HCF 오차 보고서](COMSOL_Install/waveoptics_m8/HCF_VALIDATION_REPORT.md) · [2코어 MCF 검증](COMSOL_Install/waveoptics_m8/MCF_2CORE_VALIDATION_REPORT.md) · [실행 가능한 Colab](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/COMSOL_Install/waveoptics_m8/M8_HCF_MCF_2Core_Validation.ipynb). 곡선 P2 FEM·독립 기준·메시/영역 수렴을 추가했습니다. HCF는 이상적 모세관 모델 검증입니다.

# Python Wave Optics — M1–M8

## 입력값·출력값·예시·Mesh 비교 — 통합 Colab 보고서

[▶ G.654.E / G.657.A1 입력·출력 설명서](COMSOL_Install/fiber_simulation_guide/readme.md) · [Colab 실행](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/COMSOL_Install/fiber_simulation_guide/Fiber_Input_Output_Examples.ipynb) · [실행 결과 ZIP](COMSOL_Install/fiber_simulation_guide/Fiber_Input_Output_Results.zip)

직접 Δn 입력과 MFD 보정 모드를 구분하고, 두 광섬유의 **실제 계산 예시 6개**, 분석 파장 변경, 코어 반경 변경, Mesh 수·외부 계산영역의 영향을 설명합니다. 사용자 Mesh·파장 입력을 바꾼 실제 재실행도 확인했습니다. [G.654.E 입력/출력](validation_report/LS_PremierLinkLE125_G.654.E/readme.md) · [G.657.A1 입력/출력](g657a1_validation/readme.md)


## G.657.A1 Google Colab 실행 보고서 (2026-10-19)

[▶ Colab에서 G.657.A1 FEM 실행](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/g657a1_validation/G657A1_Colab_Validation.ipynb) · [검증 보고서](g657a1_validation/G657A1_VALIDATION_REPORT.md) · [그래프·표 결과 ZIP](g657a1_validation/G657A1_Colab_Results.zip)

1310 nm MFD 9.2 µm를 참고한 가상 단면의 radial/2D FEM·분산·영역/메시 수렴을 재계산합니다. **수치 검사 11/12 통과, 기존 2D 격자의 n_eff 기준 1건 미달. 규격 수치 참고 4/4 범위 내이며 굽힘 손실·실제 제품·ITU 적합성을 검증한 것은 아닙니다.** 저장된 그림 8개와 실행 결과가 포함돼 있습니다.


## G.654.E Google Colab 실행 보고서 (2026-10-19)

[▶ Colab에서 G.654.E FEM 시뮬레이션 직접 실행](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/validation_report/LS_PremierLinkLE125_G.654.E/G654E_Colab_Validation.ipynb) · [실행 결과가 포함된 ipynb](validation_report/LS_PremierLinkLE125_G.654.E/G654E_Colab_Validation.ipynb) · [계산 그래프/표 ZIP](validation_report/LS_PremierLinkLE125_G.654.E/G654E_Colab_Results.zip)

코어 반경·파장·MFD 설정 후 **Bessel 독립 해석해, radial FEM, 기존 M7 2D FEM, 파장별 분산, 메시 수렴성**을 Colab에서 재계산하며, 기본 조건으로 실행한 그래프도 노트북에 저장했습니다. **기록된 수치 검사 11/12 통과; 2D 메시 기준 1건 미달. 실제 LE125 제품 내부 굴절률 프로파일 검증 아님.**


## G.654.E 검증 (2026-10-19)

[▶ PremierLink LE125 G.654.E 기준 수치 검증 보고서·재현 코드·그래프](validation_report/LS_PremierLinkLE125_G.654.E/README.md) — 이상적 Step-index FEM/Bessel 일치성 11/12 검사 통과, 2D FEM 메시 수렴 문제 1건 확인. **실제 LE125 굴절률 프로파일·감쇠·제품 적합성 검증 아님.**


## 2026-10-19 연구용 수치 검증 업데이트

[연구용 실행·검증 안내](research/README.md) · [메시/영역/분산 검증 결과](research/results/BENCHMARK_REPORT.md)

원형 LP01 FEM, 정확한 장 적분, 무한구간 Bessel 기준 적분, MCF 수렴 검사와 회귀 검사를 추가했습니다. 검증 범위와 미검증 범위는 위 안내를 확인하십시오.


## M4 3D 전체 구현 보고서와 첨부사진 검증 추가

[2026-10-11 결과 보고서 및 재현 소스](COMSOL_Install/reports/M4_3D_2026-10-11/)를 추가했습니다. 3D Nédélec·PML·직사각형 포트·S 행렬 구현의 기존 104개 실제 시험과 사진 기준 추가 검사 9개를 구분해 수록했습니다. 기존 아래 `waveoptics_m4` 공동 재구현은 유지하며 전체 주파수 응답 구현은 새 보고서 폴더의 소스 ZIP으로 제공합니다. 공동 주파수 0.1% 기준은 추가 세분화 메시에서 충족했으며 기존 기본 메시를 통과로 바꾼 것은 아닙니다. 3D 구 Mie 산란·임의 구조 전체 스퓨리어스 배제·사람 또는 별도 모델의 독립 검토는 미검증입니다.

> **두 가지 사용 방식 지원:** ① Python 소스를 직접 불러와 FEM 연구 및 파라미터 수정 ② 설치 없이 웹에서 광모드·분산·MCF/HCF 그래프 확인.
>
> **범위 주의:** Python M7은 약유도(weak-guidance) 2D 스칼라 FEM입니다. 웹은 Python FEM을 실행하지 않으며, LP01 독립 해석식, M8 저장된 FEM 결과의 보간, 이상적 HCF capillary 근사식을 사용합니다. 실제 COMSOL full-vector 기능·상용 광섬유 스펙 검증과 동일하다고 주장하지 않습니다.

## ① 웹페이지에서 파라미터 입력하고 그래프 보기

### **[▶ Wave Optics Lab · 웹 시뮬레이터 실행](https://kimheeseo.github.io/LSCNS/COMSOL/)**

- **M7 · Step-index SMF / G.654.E 개념:** 파장 \`λ\`, 코어 반경 \`a\`, 코어/클래딩 굴절률차 \`Δn\` 입력 → LP01 Bessel 해석식 \`n_eff\`, second-moment MFD, \`Aeff\`, 모드 분포, 파장별 \`n_eff\` 그래프. 입력된 형상은 임의의 원형 step-index이며 특정 제조사 G.654.E의 실제 프로파일을 재현한 것이 아닙니다.
- **M8 · MCF:** 검증된 M8 2코어 FEM 참조점(12/16/20/24 µm)을 코어 간격 기준으로 로그 보간해 \`Δn_eff\`, 결합 길이 및 그래프 표시. **웹에서 새 FEM을 계산하지 않습니다.**
- **M8 · HCF:** 공기 코어 반경·파장·공기 굴절률 → *ideal hollow capillary* 근사 \`n_eff\`, \`D\` 및 개념 모드 그림. NANF/ARF 반공진·PML·confinement loss는 계산하지 않습니다.
- **결과 저장:** 메인 그래프 PNG 및 CSV 다운로드.
- **원본 웹 소스:** [COMSOL_Install/index.html](COMSOL_Install/index.html) · [web/app.js](COMSOL_Install/web/app.js) · [web/style.css](COMSOL_Install/web/style.css)

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

- `COMSOL/COMSOL_Install/waveoptics_m7/screenshots/` → `m7_mode_and_convergence.png`, `m7_dispersion.png`, `m7_dispersion.csv`, `m7_summary.json`
- `COMSOL/COMSOL_Install/waveoptics_m8/screenshots/` → `m8_mcf_supermodes.png`, `m8_coupling_and_hcf.png`, `m8_summary.json`
- `COMSOL/COMSOL_Install/waveoptics_m7/M7_VALIDATION_REPORT.md`, `COMSOL/COMSOL_Install/waveoptics_m8/M8_VALIDATION_REPORT.md`
- `COMSOL/COMSOL_Install/M7_M8_EXECUTION_SUMMARY.json`, `COMSOL/COMSOL_Install/M1_M8_FINAL_REPORT.md`

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
sys.path.insert(0, str(Path("COMSOL/COMSOL_Install/waveoptics_m7/src").resolve()))

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
sys.path.insert(0, str(Path("COMSOL/COMSOL_Install/waveoptics_m8/src").resolve()))
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
| M1 | Slab scalar waveguide | [waveoptics_m1](COMSOL_Install/waveoptics_m1/) |
| M2 | Vector-mode baseline | [waveoptics_m2](COMSOL_Install/waveoptics_m2/) |
| M3 | TE/TM 2D scattering · PML · S parameters | [M3](COMSOL_Install/waveoptics_m3/) · [검증 보고서](COMSOL_Install/waveoptics_m3/M3_VALIDATION_REPORT.md) |
| M4 | 3D Nédélec PEC cavity 부분 재구현 | [M4](COMSOL_Install/waveoptics_m4/) · [검증 보고서](COMSOL_Install/waveoptics_m4/M4_VALIDATION_REPORT.md) |
| M5 | Material dispersion · spectral S (1D) | [waveoptics_m5](COMSOL_Install/waveoptics_m5/) |
| M6 | Coating inverse design (1D) | [waveoptics_m6](COMSOL_Install/waveoptics_m6/) |
| **M7** | Step-index optical fiber LP01 scalar FEM, MFD, Aeff | [M7](COMSOL_Install/waveoptics_m7/) · [검증 보고서](COMSOL_Install/waveoptics_m7/M7_VALIDATION_REPORT.md) |
| **M8** | Two-core MCF scalar FEM + HCF capillary reference | [M8](COMSOL_Install/waveoptics_m8/) · [검증 보고서](COMSOL_Install/waveoptics_m8/M8_VALIDATION_REPORT.md) |

M1~M3은 원본 소스와 수치 검증을 제공합니다. M4는 `waveoptics_m4`의 PEC 공동 솔버와 `COMSOL_Install/reports/M4_3D_2026-10-11`의 전체 3D PML·포트 구현 보관 ZIP을 구분합니다. M5·M6은 1D 독립 해석 오라클입니다. MATLAB/COMSOL 소프트웨어 자체의 검증서나 상용 광섬유 제품 사양서와의 정량 비교는 아닙니다.

## 자동 테스트

- [M7·M8 실제 Python FEM 수치 검증](https://github.com/kimheeseo/LSCNS/actions/workflows/waveoptics-m7-m8-validation.yml)
- [웹 해석식·화면 smoke test](https://github.com/kimheeseo/LSCNS/actions/workflows/waveoptics-web-smoke.yml)
- [M1–M8 통합 검증 현황](COMSOL_Install/M1_M8_FINAL_REPORT.md)

웹과 Python 계산 엔진은 명확히 분리돼 있습니다. 새 제품 사양을 비교할 때는 **Python FEM의 메시 수렴·해석해 오차**를 확인한 결과를 연구용 근거로 사용하고, 웹 미리보기 결과를 FEM 검증값으로 보고하지 마십시오.
