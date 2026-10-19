"""Generate Korean reports from actual numerical outputs, never typed-in results."""
from pathlib import Path
import json
ROOT=Path(__file__).resolve().parents[1]


def fmt(x):return f'{x:.10g}' if isinstance(x,(int,float)) else str(x)


def table(rows,cols):
    labels=[v for _,v in cols]
    return '| '+' | '.join(labels)+' |\n| '+' | '.join(['---']*len(cols))+' |\n'+'\n'.join('| '+' | '.join(fmt(r[k]) for k,_ in cols)+' |' for r in rows)+'\n'


def make_report(results=None):
    out=Path(results) if results else ROOT/'validation'/'results'
    s=json.loads((out/'summary.json').read_text());h=s['hcf_finest'];hr=s['hcf_reference'];m=s['mcf_finest'];mr=s['mcf_reference'];cfg=s['config']
    common='''오차율은 `100×|계산값−기준값|/|기준값|`이며, 유효굴절률 절대 오차는 `|Δneff|`입니다. **실제 제품·COMSOL·실측 데이터 대비 오차율은 아직 산정하지 않았습니다.** 오차율이 작아도 모델의 생략된 물리가 검증되는 것은 아닙니다.
'''
    source='''## 이론 자료

- Marcatili & Schmeltzer (1964), *Hollow Metallic and Dielectric Waveguides for Long Distance Optical Transmission and Lasers*, DOI [10.1002/j.1538-7305.1964.tb04108.x](https://doi.org/10.1002/j.1538-7305.1964.tb04108.x), [Bell Labs 공개 설명](https://www.nokia.com/bell-labs/publications-and-media/publications/hollow-metallic-and-dielectric-waveguides-for-long-distance-optical-transmission-and-lasers/).
- White et al. (2002), *Multipole method for microstructured optical fibers. I. Formulation*, DOI [10.1364/JOSAB.19.002322](https://doi.org/10.1364/JOSAB.19.002322). 원통 조화함수로 경계를 정합하는 방법 계열의 참고문헌입니다. 이번 독립 기준 코드는 직접 유도한 **스칼라 2원 경계 정합**이며, 논문의 전체 벡터 솔버나 논문 표를 재현했다고 주장하지 않습니다.

확인일: 2026-10-20 한국시간. 문헌의 일반적 방법과 이번 구현·검증 범위를 구분했습니다.
'''
    hcf=f'''# HCF 모세관 모델 검증 보고서 — M8 개선

검증 대상은 **이상적 원형 중공 모세관의 스칼라 Dirichlet 모델**입니다. ARF/NANF 실물 광섬유의 성능 검증 보고서와 구분해야 합니다.

{common}
## 조건과 결과

| 항목 | 값 |
|---|---:|
| 파장 | {cfg['wavelength_um']} µm |
| 중공 반경 | {cfg['hcf_radius_um']} µm |
| 공기 굴절률 | {cfg['n_air']} (파장에 대해 일정) |
| 최종 방사형 P1 요소 | {h['elements']} |
| neff 해석식 | {hr['n_eff']:.13f} |
| neff FEM | {h['n_eff']:.13f} |
| neff 절대 오차 | {h['neff_abs_error']:.6g} |
| neff 상대 오차 | {h['neff_error_pct']:.6g}% |
| D 해석식 | {hr['D_ps_nm_km']:.10f} ps/(nm·km) |
| D FEM | {h['D_ps_nm_km']:.10f} ps/(nm·km) |
| D 상대 오차 | {h['D_error_pct']:.6g}% |
| 횡방향 고유값 kt² 상대 오차 | {h['kt2_error_pct']:.6g}% |
| Aeff 해석적 장 적분 | {hr['Aeff_um2']:.9f} µm² |
| Aeff FEM 장 적분 | {h['Aeff_um2']:.9f} µm² |
| Aeff 상대 오차 | {h['Aeff_error_pct']:.6g}% |

neff는 약 1이므로 상대 오차가 작게 보입니다. 파동의 횡방향 분포를 결정하는 **kt² 오차도 함께 확인**했습니다.

## 독립 검증 방법

기준 해는 `J0(u01)=0`인 첫 근으로부터 `kt²=(u01/R)²`, `neff=sqrt(n_air²−kt²·λ²/(2π)²)`를 계산합니다.
FEM은 이 근을 입력하지 않고 `−(rψ′)′=kt²rψ`, 중심에서 자연 정칙 조건, 외벽에서 `ψ(R)=0`을 이산화해 고유값과 장을 구합니다.

`q=kt²/(2π)²`일 때 일정한 공기 굴절률을 가정한 분산은 `D=(λ/c)·10¹²·q·n_air²/neff³`입니다. 여기서 λ는 µm입니다. 수치 FEM의 서로 다른 5개 파장 결과에 대한 5점 미분도 별도 검사했습니다.

### 메시 수렴

{table(s['hcf_mesh'],[('elements','요소'),('neff_abs_error','neff 절대 오차'),('kt2_error_pct','kt² 오차 %'),('D_error_pct','D 오차 %'),('Aeff_error_pct','Aeff 오차 %')])}

오차는 P1 공간 이산화에서 발생하며 세분화에 따라 줄어듭니다. 이 문제의 R은 실제 모델 반경이므로 R 변경을 메시 수렴으로 해석하면 안 됩니다.

### 반경·파장 9개 조합

{table(s['hcf_sweep'],[('radius_um','R µm'),('wavelength_um','λ µm'),('neff_abs_error','neff 절대 오차'),('D_error_pct','D 오차 %')])}

### 분산 미분 간격 검사

{table(s['hcf_derivative'],[('step_um','미분 간격 µm'),('D_FEM_fivepoint','5점 D'),('error_vs_analytic_pct','기준 D 대비 오차 %')])}

## 검증되지 않은 항목

- 유리벽 두께·튜브 수·튜브 간격·중첩 튜브·반공진 대역과 공진 피크.
- 복소 유효굴절률, PML, 구속 손실(dB/km), 굽힘 손실, 표면 산란, 기체 물질 분산.
- 실제 HCF 측정값 또는 독립 full-vector Maxwell 계산과의 일치.

현재 코드로 실제 HCF 손실을 출력하거나, 위 오차율을 상용 HCF 정확도로 해석하면 안 됩니다. 다음 개발에는 구체적인 HCF 단면과 재료·실측 또는 공개 벡터 기준 데이터가 필요합니다.

[원시 데이터](validation/results/hcf_mesh.csv) · [Colab 보고서](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/waveoptics_m8/M8_HCF_MCF_2Core_Validation.ipynb)

{source}'''
    mcf=f'''# 2코어 MCF 검증 보고서 — M8 개선

**동일한 원형 스텝형 코어 2개, 직선·무손실·약도파 스칼라 모델**의 짝/홀 슈퍼모드와 결합 길이를 검증했습니다.

{common}
## 입력과 기본 결과

| 입력 | 값 |
|---|---:|
| 파장 | {cfg['wavelength_um']} µm |
| 코어 반경 | {cfg['core_radius_um']} µm |
| 코어−클래딩 Δn | {cfg['delta_n']} |
| 코어 중심 간 거리 | {cfg['pitch_um']} µm |
| 클래딩 굴절률 | 실리카 Sellmeier |
| 최종 영역 | −{cfg['halfwidth_um']}…+{cfg['halfwidth_um']} µm 정사각형 |
| 최종 목표 메시 크기 | {m['mesh_size_um']} µm |
| 1/4 영역 요소 / P2 자유도 | {m['quarter_elements']} / {m['quarter_dofs']} |

| 출력 | 독립 무한영역 기준 | 개선 FEM | 상대 오차 % |
|---|---:|---:|---:|
| 짝모드 neff | {mr['n_even']:.13f} | {m['n_even']:.13f} | {100*m['neff_even_abs_error']/abs(mr['n_even']):.6g} |
| 홀모드 neff | {mr['n_odd']:.13f} | {m['n_odd']:.13f} | {100*m['neff_odd_abs_error']/abs(mr['n_odd']):.6g} |
| Δneff | {mr['index_split']:.12g} | {m['index_split']:.12g} | {m['split_error_pct']:.6g} |
| 결합 길이 Lc (mm) | {mr['coupling_length_mm']:.10f} | {m['coupling_length_mm']:.10f} | {m['Lc_error_pct']:.6g} |

짝/홀 neff의 최대 절대 오차는 {max(m['neff_even_abs_error'],m['neff_odd_abs_error']):.6g}입니다. neff 자체 오차가 작아도 두 값의 차이인 Δneff는 상대 오차가 커질 수 있으므로 Δneff와 Lc를 직접 판정합니다.

## 개선 사항과 독립 기준

기존 M7 기반 Cartesian P1은 요소 중심으로 원형 코어 재료를 분류했습니다. 메시가 바뀔 때 경계 모양도 달라져 결과가 비단조적으로 변합니다.
개선 모델은 Gmsh의 원호 경계, 원 위에 놓인 P2 중간 노드, 곡선 기하 매핑, 코어/클래딩별 재료를 사용합니다. 원은 유한 차수 기하로 근사하므로 여전히 메시 검증이 필요합니다.

1/4 영역에서 y=0은 Neumann 대칭 조건입니다. x=0은 짝모드 Neumann, 홀모드 Dirichlet입니다. 외곽 x=L, y=L은 Dirichlet입니다. 동일 메시에서 두 대칭 문제를 풀어 메시 비대칭에 의한 인공 분리를 피합니다.

독립 기준은 유한영역이나 FEM을 사용하지 않습니다. 각 코어 내부 `J_m(qr)cos(mθ)`, 외부 `K_m(wr)cos(mθ)`를 전개하고 원형 경계에서 장과 법선 미분을 정합합니다. x 반사 대칭은 `b_right,m=±(−1)^m b_left,m`입니다. 무한 클래딩에서 감쇠하는 해를 직접 사용하며 차수·각도 샘플 수·분리 코어 극한을 검사했습니다.

{table(s['mcf_reference_orders'],[('order','원통함수 차수'),('index_split','Δneff'),('coupling_length_mm','Lc mm')])}

본 보고서의 경계 정합 기준은 동일 스칼라 방정식에 대한 독립 수치 해입니다. 유한 차수 오차는 남아 있으며, 차수 수렴 검사를 통해 작음을 확인했습니다.

## 메시 수렴: 영역 고정

{table(s['mcf_mesh'],[('mesh_size_um','목표 h µm'),('quarter_dofs','1/4 자유도'),('neff_even_abs_error','짝 neff 절대 오차'),('split_error_pct','Δneff 오차 %'),('Lc_error_pct','Lc 오차 %')])}

## 영역 수렴: 목표 h 고정

{table(s['mcf_domain'],[('halfwidth_um','영역 반폭 µm'),('quarter_dofs','1/4 자유도'),('coupling_length_mm','Lc mm'),('Lc_error_pct','Lc 오차 %')])}

영역별로 재메시하므로 이 표는 잔여 메시 오차도 포함합니다. 도메인 변화와 메시 세분화를 별도로 검사해야 합니다.

## 중심 간격별 검증

{table(s['mcf_pitch'],[('pitch_um','간격 µm'),('reference_Lc_mm','기준 Lc mm'),('coupling_length_mm','FEM Lc mm'),('Lc_error_pct','Lc 오차 %')])}

## 기존 P1과 비교

{table(s['mcf_legacy'],[('n_grid','기존 grid'),('coupling_length_mm','기존 Lc mm'),('split_error_pct','Δneff 오차 %'),('Lc_error_pct','Lc 오차 %')]) if s['mcf_legacy'] else '이 실행에서 기존 P1 비교를 생략했습니다.'}

기존 P1의 영역 반폭은 27 µm입니다. 개선 비교는 메시·기하 차수·재료 경계·대칭 처리·영역 크기를 함께 바꾼 전체 변경 결과입니다. 개선 폭을 메시 차수 하나의 효과라고 단정하지 않습니다.

## 전력 전달과 사용 범위

`κ=πΔneff/λ`, `Lc=λ/(2Δneff)=π/(2κ)`이며, 길이 단위 변환 후 `P2(z)=sin²(πz/(2Lc))`, `P1=1−P2`입니다.
여기서 Lc는 처음으로 전력이 완전히 전달되는 길이입니다. 전체 비팅 주기는 2Lc입니다.

이 전력식은 동일 코어의 두 슈퍼모드를 사용한 코어 국소 상태 모델입니다. 실제 입력장 중첩·고차모드·방사모드까지 계산한 launch 효율은 아닙니다.
굽힘, 비틀림, 제조 불균일, 편광 결합과 길이별 랜덤 결합을 포함하지 않으므로 상용 MCF의 케이블 XT(dB/km)로 사용하면 안 됩니다.
독립 기준은 단일 코어 V<2.405, 동일 반경·동일 Δn, 서로 겹치지 않는 두 코어에 제한됩니다. 새로운 조건은 메시·영역·기준 차수 수렴을 다시 검사해야 합니다.

![실제 FEM 슈퍼모드](validation/results/02_mcf_supermodes.png)
![결합 길이와 전력 전달](validation/results/03_coupling.png)

{source}'''
    (ROOT/'HCF_VALIDATION_REPORT.md').write_text(hcf)
    (ROOT/'MCF_2CORE_VALIDATION_REPORT.md').write_text(mcf)
    overview=f'''# M8 — HCF / 2코어 MCF 검증 개선 보고서

검증 판정 **{s['passed']}/{s['total']} 통과**. 테스트 결과는 [tests.xml](validation/results/tests.xml), [원시 판정표](validation/results/checks.csv)를 확인하십시오. 테스트 통과 개수는 물리적 오차율이 아닙니다.

| 대상 | 비교 기준 | 기본 조건의 오차 |
|---|---|---:|
| HCF 모세관 neff | 이상적 모세관 Bessel 해 | 절대 {h['neff_abs_error']:.6g}, 상대 {h['neff_error_pct']:.6g}% |
| HCF 모세관 D | 이상적 모세관 해석 미분 | {h['D_error_pct']:.6g}% |
| 2코어 MCF Δneff | 무한영역 원통함수 경계 정합 | {m['split_error_pct']:.6g}% |
| 2코어 MCF Lc | 무한영역 원통함수 경계 정합 | {m['Lc_error_pct']:.6g}% |

{common}

- [HCF 상세 보고서](HCF_VALIDATION_REPORT.md)
- [2코어 MCF 상세 보고서](MCF_2CORE_VALIDATION_REPORT.md)
- [Colab: 실행 결과 포함](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/waveoptics_m8/M8_HCF_MCF_2Core_Validation.ipynb)
- [입력·실행 방법](README.md)

![메시와 영역 수렴](validation/results/01_convergence.png)

기존 `screenshots/`와 Cartesian `solve_two_core`는 과거 결과·호환용으로 유지합니다. 현재 검증 결과는 `validation/results/`이며 서로 다른 결과를 섞어 해석하지 마십시오.
'''
    (ROOT/'M8_VALIDATION_REPORT.md').write_text(overview)
    readme=f'''# M8: HCF 모세관 · 2코어 MCF 연구용 검증

이 도구는 **이상적 HCF 모세관의 스칼라 모드**와 **동일 원형 코어 2개의 스칼라 결합 모드**를 계산합니다. 실제 ARF/NANF 손실 또는 실제 MCF 랜덤 누화 검증까지 완료된 도구는 아닙니다.

[통합 결과 보고서](M8_VALIDATION_REPORT.md) · [HCF 보고서](HCF_VALIDATION_REPORT.md) · [2코어 MCF 보고서](MCF_2CORE_VALIDATION_REPORT.md)

[▶ Colab에서 실행](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/waveoptics_m8/M8_HCF_MCF_2Core_Validation.ipynb)

Colab에서 런타임 연결 후 **런타임 → 모두 실행**을 선택하십시오. `CONFIG`에서 입력을 바꾸면 재계산합니다. 엔진은 노트북에 내장되어 저장소 복제 없이 실행할 수 있습니다. Python 환경에서 내장 셀 재실행을 검증했으며, Colab 브라우저 UI의 수동 실행은 확인하지 않았습니다.

## 어떤 입력으로 무엇을 계산하는가

| 모델 | 주요 입력 (기본값) | 출력 |
|---|---|---|
| HCF 모세관 | λ={cfg['wavelength_um']} µm, 반경={cfg['hcf_radius_um']} µm, 일정 n_air={cfg['n_air']}, 요소 수 | neff, 횡방향 kt², D [ps/(nm·km)], Aeff [µm²], 고유값 잔차·메시 오차 |
| 2코어 MCF | λ={cfg['wavelength_um']} µm, 반경={cfg['core_radius_um']} µm, Δn={cfg['delta_n']}, 중심 간격={cfg['pitch_um']} µm | 짝/홀 neff, Δneff, κ [1/m], Lc [mm], 슈퍼모드 장 |
| MCF 검증 | 목표 h={cfg['mesh_sizes_um']} µm, 영역 반폭={cfg['domain_halfwidths_um']} µm | 기준 대비 오차·메시/영역 수렴·간격별 비교 |
| MCF 전력 모델 | 전파 길이 z [m], Lc [mm] | 무손실 동일 코어의 P1/P2 전력 분율 |

예: 기본 2코어 조건에서 개선 FEM Lc={m['coupling_length_mm']:.8f} mm입니다. 전력 모델에서는 z=Lc에서 P2=1, z=2Lc에서 P2=0입니다. 이 수치는 실제 케이블 XT 스펙을 뜻하지 않습니다.

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

원시 결과는 `validation/results/`에 CSV·JSON·PNG로 저장됩니다. 판정 실패 시 CLI는 오류 종료하며, 실패 수치를 저장합니다. 노트북은 실패 항목도 화면에 표시합니다. 전체 실행 시간은 이 환경에서 약 {s['runtime_seconds']:.0f}초이며 하드웨어와 메시 크기에 따라 달라집니다.

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

{common}

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
'''
    (ROOT/'README.md').write_text(readme)
    (ROOT/'readme.md').write_text(readme)
    return s


if __name__=='__main__':make_report()
