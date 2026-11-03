"""Execute independent analytical comparisons and write the M1 report.

Run pytest first. All tables below are built from this execution, not copied
from prior solver outputs. The solver never imports tests.analytic_slab.
"""
import csv
from datetime import datetime, timezone
import importlib.metadata
import json
from pathlib import Path
import platform
import hashlib
import xml.etree.ElementTree as ET
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from waveoptics import solve_slab
from tests.analytic_slab import exact_modes, l2_field_error

OUT = Path(__file__).resolve().parent/'results'
OUT.mkdir(exist_ok=True)
BASE = dict(n_core=1.5,n_clad=1.45,width_um=3.0,wavelength_um=1.55,padding_um=20.0,max_modes=6)
STRONG = dict(n_core=3.45,n_clad=1.44,width_um=0.4,wavelength_um=1.55,padding_um=8.0,max_modes=6)


def save_csv(filename, rows):
    with (OUT/filename).open('w',newline='') as f:
        writer = csv.DictWriter(f,fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def mdtable(headers, rows):
    return '\n'.join(['| '+' | '.join(headers)+' |','| '+' | '.join(['---']*len(headers))+' |']+
                     ['| '+' | '.join(map(str,r))+' |' for r in rows])


def study(parameters, hs, name):
    rows, solutions = [], {}
    for pol in ('TE','TM'):
        refs = exact_modes(parameters['n_core'],parameters['n_clad'],parameters['width_um'],parameters['wavelength_um'],pol)
        previous = {}
        for level,h in enumerate(hs):
            sol = solve_slab(**parameters,h_um=h,polarization=pol)
            solutions[(pol,level)] = sol
            if len(sol.modes) != len(refs):
                raise AssertionError('Analytical/FEM mode counts differ.')
            for mode,ref in zip(sol.modes,refs):
                error = abs(mode.neff-ref.neff)
                field_error = l2_field_error(sol,mode,ref)
                prev = previous.get(mode.order)
                row = {'case':name,'mode':f'{pol}{mode.order}','level':level+1,
                       'h_max_um':sol.h_max_um,'elements':sol.elements,'dofs':sol.dofs,
                       'neff_exact':ref.neff,'neff_fem':mode.neff,'abs_neff_error':error,
                       'neff_order':float(np.log(prev[1]/error)/np.log(prev[0]/sol.h_max_um)) if prev else None,
                       'relative_l2_field_error':field_error,
                       'field_order':float(np.log(prev[2]/field_error)/np.log(prev[0]/sol.h_max_um)) if prev else None,
                       'relative_residual':mode.relative_residual,'decay_estimate':mode.boundary_decay_estimate}
                rows.append(row)
                previous[mode.order]=(sol.h_max_um,error,field_error)
    save_csv(name+'_convergence.csv',rows)
    # Show FE fields and independently computed exact fields.
    fig,axes = plt.subplots(2,2,figsize=(11,6),sharex=True)
    for pol,axisrow in zip(('TE','TM'),axes):
        sol = solutions[(pol,len(hs)-1)]
        refs = exact_modes(parameters['n_core'],parameters['n_clad'],parameters['width_um'],parameters['wavelength_um'],pol)
        for j,ax in enumerate(axisrow):
            mode,ref = sol.modes[j],refs[j]
            xp = np.linspace(sol.x_um[0],sol.x_um[-1],6001)
            fp = np.interp(xp,sol.x_um,mode.field)
            exact = ref.field(xp)
            if np.dot(fp,exact) < 0:
                fp = -fp
            ax.plot(xp,exact,'--',lw=2,color='#e38b2c',label='infinite-cladding analytical')
            ax.plot(xp,fp,lw=1,color='#147b9d',label='gmsh + P1 FEM')
            a = parameters['width_um']/2
            ax.axvspan(-a,a,color='#d7e9ef',alpha=0.45)
            view = 6 if name=='baseline' else 1.5
            ax.set(xlim=(-view,view),xlabel='x (um)',ylabel=('Ey' if pol=='TE' else 'Hy')+' | L2 normalized',
                   title=f'{pol}{j}: n_eff={mode.neff:.9f}')
            ax.grid(alpha=0.2)
            ax.legend(fontsize=7)
        data = np.column_stack([sol.x_um]+[m.field for m in sol.modes])
        np.savetxt(OUT/f'{name}_fields_{pol}.csv',data,delimiter=',',
                   header='x_um,'+','.join(f'{pol}{m.order}' for m in sol.modes),comments='')
    fig.suptitle(f"{name}: nc={parameters['n_core']}, ns={parameters['n_clad']}, width={parameters['width_um']} um, wavelength=1.55 um")
    fig.tight_layout()
    fig.savefig(OUT/f'{name}_fields.png',dpi=180)
    plt.close(fig)
    return rows


baseline = study(BASE,[0.15,0.075,0.0375],'baseline')
strong = study(STRONG,[0.01,0.005,0.0025],'high_contrast')
boundary = []
for name,params,pads,h in [('baseline',BASE,[15,20,25],0.05),('high_contrast',STRONG,[4,6,8],0.0025)]:
    for pol in ('TE','TM'):
        prev = None
        for pad in pads:
            sol = solve_slab(**{**params,'padding_um':pad},polarization=pol,h_um=h)
            for mode in sol.modes:
                boundary.append({'case':name,'mode':f'{pol}{mode.order}','padding_um':pad,
                                 'h_max_um':sol.h_max_um,'neff_fem':mode.neff,
                                 'delta_from_previous_padding':abs(mode.neff-prev.modes[mode.order].neff) if prev else None,
                                 'decay_estimate':mode.boundary_decay_estimate})
            prev = sol
save_csv('boundary_convergence.csv',boundary)

flux = []
for pol in ('TE','TM'):
    prev = None
    for h in [0.01,0.005,0.0025]:
        sol = solve_slab(**STRONG,polarization=pol,h_um=h)
        j = np.argmin(abs(sol.x_um-0.2))
        f = sol.modes[0].field
        left = (f[j]-f[j-1])/(sol.x_um[j]-sol.x_um[j-1])
        right = (f[j+1]-f[j])/(sol.x_um[j+1]-sol.x_um[j])
        if pol=='TM':
            left /= STRONG['n_core']**2
            right /= STRONG['n_clad']**2
        err = abs(left-right)/max(abs(left),abs(right))
        flux.append({'mode':pol+'0','h_max_um':sol.h_max_um,'relative_flux_jump':err,
                     'order':float(np.log(prev/err)/np.log(2)) if prev else None})
        prev = err
save_csv('interface_flux.csv',flux)

fig,axes = plt.subplots(1,2,figsize=(10,4.1))
colors = ['#147b9d','#cc6733','#3f915c','#8b5eb4']
for name,rows,ax in [('Baseline',baseline,axes[0]),('High contrast',strong,axes[1])]:
    for mode,color in zip(['TE0','TE1','TM0','TM1'],colors):
        selected = [r for r in rows if r['mode']==mode]
        ax.loglog([r['h_max_um'] for r in selected],[r['abs_neff_error'] for r in selected],'o-',color=color,label=mode)
    h0 = rows[0]['h_max_um']
    hs = np.array(sorted(set(r['h_max_um'] for r in rows)))
    ax.loglog(hs,rows[0]['abs_neff_error']*(hs/h0)**2,'k--',alpha=0.6,label='slope 2 reference')
    ax.set(xlabel='maximum element size h (um)',ylabel='absolute n_eff error',title=name)
    ax.legend(fontsize=8)
    ax.grid(which='both',alpha=0.2)
fig.tight_layout()
fig.savefig(OUT/'convergence.png',dpi=180)
plt.close(fig)

testtree = ET.parse(OUT/'pytest.xml').getroot()
tests = testtree.findall('.//testcase')
passed = sum(not any(t.find(tag) is not None for tag in ['failure','error','skipped']) for t in tests)
assert passed == len(tests) == 32
metadata = {'executed_utc':datetime.now(timezone.utc).isoformat(),
            'python':platform.python_version(),'platform':platform.platform(),
            'packages':{name:importlib.metadata.version(name) for name in ['numpy','scipy','scikit-fem','gmsh','matplotlib','pytest']},
            'pytest_passed':passed,'pytest_total':len(tests),
            'baseline':BASE,'high_contrast':STRONG,
            'test_hashes':{str(p.relative_to(Path(__file__).parent)):hashlib.sha256(p.read_bytes()).hexdigest()
                           for p in [Path(__file__).parent/'tests/analytic_slab.py',Path(__file__).parent/'tests/test_m1.py']}}
allresults = {'metadata':metadata,'baseline':baseline,'high_contrast':strong,'boundary':boundary,'interface_flux':flux}
(OUT/'validation.json').write_text(json.dumps(allresults,indent=2)+'\n')


def convergence_table(rows):
    return mdtable(['모드','h_max (μm)','요소 / 자유도','해석 n_eff','FEM n_eff','절대오차','n_eff 차수','필드 상대 L2 오차','필드 차수'],[
        [r['mode'],f"{r['h_max_um']:.7f}",f"{r['elements']} / {r['dofs']}",f"{r['neff_exact']:.12f}",
         f"{r['neff_fem']:.12f}",f"{r['abs_neff_error']:.6e}",f"{r['neff_order']:.5f}" if r['neff_order'] else '—',
         f"{r['relative_l2_field_error']:.6e}",f"{r['field_order']:.5f}" if r['field_order'] else '—'] for r in rows])


report = r'''# M1 검증 보고서: 1D 슬랩 TE/TM FEM 모드 솔버

M1만 구현했습니다. 본 결과는 COMSOL과의 직접 비교 결과가 아니라 독립 해석해에 대한 검증입니다.
실행 일시와 라이브러리 버전은 `results/validation.json`에 기록했습니다.

## 1. 수식과 이산화 — 구현 전에 설명한 정의

비자성(μ_r=1), 등방성, 실수 양의 굴절률, 무손실 매질이며 y 방향으로 무한하고 z 방향으로 균일합니다.
시간/전파 관례는 exp(iβz−iωt)입니다. 모든 길이는 μm, β는 μm⁻¹를 사용합니다.
TE에서는 ψ=E_y, TM에서는 ψ=H_y를 풉니다. k₀=2π/λ₀, n_eff=β/k₀입니다.

$$
\mathrm{TE}:\quad\psi''+k_0^2 n^2(x)\psi=\beta^2\psi,
$$
$$
\mathrm{TM}:\quad\left(n^{-2}(x)\psi'\right)'+k_0^2\psi=\beta^2 n^{-2}(x)\psi.
$$

통합식 (pψ′)′+k₀²qψ=β²rψ에서 TE는 (p,q,r)=(1,n²,1), TM은 (n⁻²,1,n⁻²)입니다.
ψ,v∈H₀¹(Ω)에 대해 부분적분한 약형식은

$$
-\int_\Omega p\,\psi'\overline{v'}\,dx+k_0^2\int_\Omega q\,\psi\bar v\,dx
=\beta^2\int_\Omega r\,\psi\bar v\,dx.
$$

계면의 ψ와 pψ′는 연속입니다. TE의 ψ′ 연속과 TM의 ψ′/n² 연속을 구분합니다.
불연속 굴절률을 노드 사이에서 선형 보간하지 않고 계면 정렬 요소별 상수로 적분합니다.
무한 클래딩은 충분한 길이 L_pad를 두고 양 끝 ψ=0으로 절단합니다. 이는 정확한 개방경계나 PML이 아닙니다.

gmsh의 세 구간 transfinite 1D 메시를 scikit-fem으로 전달해 P1 연속 Lagrange FEM을 조립합니다.
요소 길이 h_e에 대해

$$
K_e=\frac{p_e}{h_e}\begin{bmatrix}1&-1\\-1&1\end{bmatrix},\quad
Q_e=\frac{q_e h_e}{6}\begin{bmatrix}2&1\\1&2\end{bmatrix},\quad
R_e=\frac{r_e h_e}{6}\begin{bmatrix}2&1\\1&2\end{bmatrix}.
$$

경계 자유도를 제거하고 (−K+k₀²Q)c=β²Rc를 풉니다. SciPy eigsh의 shift-invert 이동점은
k₀²n_c²보다 크게 설정해 가장 큰 β²부터 찾습니다. n_s<n_eff<n_c인 모드만 반환합니다.
`max_modes`는 요청 고유쌍 수이며 전체 모드 수의 보증이 아닙니다. 후보가 모두 유도 모드이면 경고합니다.
1D에서 TE/TM 분리는 Maxwell 방정식의 정확한 축약이므로 P1 스칼라 요소를 사용할 수 있습니다.
이는 M2의 2D 벡터 문제에서 nodal 벡터 요소를 써도 된다는 의미가 아닙니다.

## 2. 독립 기준과 테스트 선작성

코어 반두께 a=width/2, u=a√(k₀²n_c²−β²), w=a√(β²−k₀²n_s²), V=a k₀√(n_c²−n_s²).
여기서 V는 **반두께 기준**이며 전두께로 정의한 자료의 V와 2배 차이납니다.

$$
u^2+w^2=V^2,\quad
u\tan u=\rho w\;(\text{짝수 모드}),\quad
-u\cot u=\rho w\;(\text{홀수 모드}),
\quad \rho_{TE}=1,\ \rho_{TM}=n_c^2/n_s^2.
$$

0부터 센 모드 m에 대해 mπ/2<u<min((m+1)π/2,V)에서
u−mπ/2−atan2(ρ√(V²−u²),u)=0을 SciPy brentq로 풉니다.
이는 FEM 결과와 무관한 해석적 분산식의 수치적 근입니다. 코어 필드는 cos(ux/a) 또는 sin(ux/a),
클래딩은 계면에서 연속인 exp(−w(|x|−a)/a)입니다. 해석 필드의 무한영역 L2 정규화 적분을 닫힌식으로 계산합니다.

해석 기준 모듈은 솔버를 import하지 않으며 솔버도 해석 기준 모듈을 import하지 않습니다.
별도의 닫힌식 특수 경우 u=π/4, w=u/ρ로 해석 기준 자체를 검사했습니다.
첫 테스트 실행에서는 해석 기준 2개 통과, 솔버 모듈 부재로 26개 실패했습니다.
이후 첫 구현에서 기존 28개가 모두 통과했고, 경계 감쇠 경고를 확인할 추가 4개 테스트를 작성했습니다.
최초 두 테스트 파일의 SHA256은 구현 전후 동일합니다. 허용오차를 완화하지 않았습니다.

최초 고정 기준: 기본 슬랩 |Δn_eff|<10⁻⁵, 고대비 슬랩 <2×10⁻⁴,
필드 상대 L2 오차 <10⁻³, 상대 고유값 잔차 <10⁻⁹,
P1의 n_eff/필드 차수 1.8~2.2, 계면 편측 기울기 플럭스 차수 0.8~1.2 및 최종 오차 <3%.
기본 슬랩 padding 15→20 μm의 Δn_eff<10⁻¹⁰.
추가 고대비 영역 기준: 4→6 μm <10⁻⁸, 6→8 μm <10⁻¹⁰.
L2 오차는 각 요소의 8점 Gauss 적분으로 실제 P1 보간 필드를 해석 필드와 비교하며 부호를 맞춥니다.

## 3. 실제 실행 환경 및 결과

ENVIRONMENT

**PYTEST_RESULT** 초기/최종 stdout, JUnit XML 및 초기 해시는 results 폴더에 포함했습니다.
경고는 실패를 숨긴 것이 아니라 작은 클래딩의 감쇠 길이에 대한 별도 진단입니다.

### 기본 슬랩: n_c=1.50, n_s=1.45, width=3.0 μm, λ₀=1.55 μm, 양쪽 padding=20 μm

BASE_TABLE

### 고굴절률 대비 슬랩: n_c=3.45, n_s=1.44, width=0.4 μm, λ₀=1.55 μm, 양쪽 padding=8 μm

굴절률은 벤치마크 상수이며 특정 제품의 측정값이나 파장분산 데이터는 아닙니다.

STRONG_TABLE

수렴 차수는 p=log(e_coarse/e_fine)/log(h_coarse/h_fine)입니다.
각 h는 실제 gmsh 메시의 최대 요소 길이입니다. 3단계 모두 전체 영역을 세분화했습니다.
예상 h²는 계면 정렬 및 충분한 클래딩에서의 P1 고유값·L2 필드 오차에 대한 것입니다.

![메시 수렴](results/convergence.png)

![기본 슬랩 필드](results/baseline_fields.png)

![고대비 슬랩 필드](results/high_contrast_fields.png)

TE 그림은 E_y, TM 그림은 H_y이며 ∫|ψ|²dx=1로 각각 정규화했습니다.
이 정규화는 SI 전계/자계 단위나 1 W 입력 정규화가 아닙니다. 두 편광의 곡선 높이를 전력으로 비교하면 안 됩니다.

### 계산영역 확대 — 메시 크기와 분리한 경계 절단 확인

BOUNDARY_TABLE

동일 h에서 padding만 확대했습니다. 감쇠 추정 exp(−αL_pad)는 경계 절단 오차의 엄밀한 상한이 아닙니다.
특히 고대비 TM1의 padding=4 μm에서 약한 감쇠 진단이 있었지만, 영역을 확대한 변화는 위 표와 같습니다.
경계 경고는 임의 입력을 위해 남겨두었습니다.

### 계면 플럭스와 고유값 잔차

FLUX_TABLE

P1은 ψ를 강하게 연속시켜도 편측 미분값을 정확히 일치시키지는 않습니다.
위 표는 고대비 기본 모드의 pψ′ 편측 근사가 1차로 연속 플럭스에 수렴하는지 검사한 결과입니다.
기본/고대비 메시 검증에서 최대 상대 고유값 잔차: MAX_RESIDUAL.
잔차는 조립된 행렬의 풀이 정확도를 나타내며 단독으로 물리 모델의 정확도를 보증하지 않습니다.

## 4. 검증된 범위

- 무손실·비자성·등방성 대칭 step-index 1D 슬랩의 TE0/TE1/TM0/TM1.
- 기본 슬랩 및 고굴절률 대비 슬랩에서 n_eff와 E_y/H_y 프로파일의 독립 해석해 일치.
- 각 슬랩에서 3단계 gmsh 메시의 n_eff 및 필드 L2 오차 2차 수렴.
- 고대비 기본 모드의 계면 플럭스 1차 수렴, 경계 자유도 제거, 스칼라 필드 L2 정규화.
- width=1.0 μm 기본 굴절률 슬랩의 단일 모드 수와 n_eff.
- 기본/고대비 영역 확대, 기본 TM의 gmsh/native 메시 결과 일치, 12개 잘못된 입력의 거부.

## 5. 검증되지 않은 범위

- 모든 굴절률/폭/파장 조합, 컷오프 직전의 매우 약하게 구속된 모드, TE2/TM2 이상.
- 비대칭/다층/연속 굴절률 분포, 물질분산, 복소 굴절률과 흡수·누설 손실.
- SI 전계/자계 복원, 전체 벡터 성분, 전력 정규화, 광섬유 MFD 및 2D 유효 모드 면적.
- COMSOL 출력과 직접 비교, 실제 제품 또는 측정 데이터와 비교.
- M2의 Nédélec 혼합 벡터 모드, M3의 PML/포트/S-파라미터, M4 3D.

## 6. 알려진 한계

- M1 API는 대칭 단일 코어 step-index만 받습니다. 임의 n(x) 입력을 제공한다고 주장하지 않습니다.
- Dirichlet 절단은 개방경계가 아닙니다. 새 입력에서는 메시 수렴과 영역 수렴을 각각 확인해야 합니다.
- 컷오프 근처 모드는 큰 클래딩이 필요하며 유한영역 때문에 누락될 수 있습니다.
- n_eff>n_s+10⁻¹⁰ 판정은 수치적 보호장치로, 극도로 컷오프에 가까운 모드를 제외할 수 있습니다.
- max_modes가 작으면 전체 모드를 찾지 못합니다. mode_limit_reached 진단 후 요청 수를 늘리십시오.
- 복소/이방성/자성 재료는 지원하지 않으며 복소 입력은 거부합니다. 손실값 0을 계산 결과처럼 반환하지 않습니다.
- gmsh는 시스템 공유 라이브러리가 필요합니다. 이 실행에서는 libXft를 별도로 로드했습니다.
  Linux 설치 의존성은 README에 설명했습니다. native 백엔드도 제공하지만 기본 검증은 실제 gmsh로 실행했습니다.
- gmsh는 전역 상태를 사용하므로 이 모듈은 이미 초기화된 gmsh 세션을 거부합니다. 같은 프로세스의 동시 gmsh 실행은 검증하지 않았습니다.

## 7. 다음 마일스톤

M1 통과 후 여기서 멈춥니다. M2 코드는 포함하지 않습니다. M2 착수 시에는 β 고유값에 대한
Maxwell 혼합 약형식, Nédélec/Lagrange 공간, 스퓨리어스 판별, 독립 원형 광섬유 벡터 해석해를 먼저 정해야 합니다.

## 기준 자료

- BYU ECE 360, §7.3 Dielectric Slab Waveguide, 식 (7.77), (7.78), (7.82), (7.91), (7.92):
  https://ece360web.groups.et.byu.net/notes/ln_dielectric_slab.pdf
- scikit-fem 공식 문서 — MeshLine, Basis, BilinearForm:
  https://scikit-fem.readthedocs.io/en/stable/api.html
- SciPy 공식 eigsh 문서 — 일반화 대칭 고유값 문제와 shift-invert:
  https://docs.scipy.org/doc/scipy-1.15.2/reference/generated/scipy.sparse.linalg.eigsh.html
'''
env = f"Python **{metadata['python']}**; "+'; '.join(f"{k} {v}" for k,v in metadata['packages'].items())+f". UTC 실행: {metadata['executed_utc']}."
boundary_table = mdtable(['조건','모드','padding (μm)','n_eff','이전 padding 대비 |Δn_eff|'],[
    [r['case'],r['mode'],r['padding_um'],f"{r['neff_fem']:.12f}",
     f"{r['delta_from_previous_padding']:.6e}" if r['delta_from_previous_padding'] is not None else '—'] for r in boundary])
flux_table = mdtable(['모드','h_max (μm)','상대 플럭스 불일치','차수'],[
    [r['mode'],f"{r['h_max_um']:.7f}",f"{r['relative_flux_jump']:.6e}",f"{r['order']:.5f}" if r['order'] else '—'] for r in flux])
for key,value in {'ENVIRONMENT':env,'PYTEST_RESULT':f'{passed}/{len(tests)} tests passed.',
                  'BASE_TABLE':convergence_table(baseline),'STRONG_TABLE':convergence_table(strong),
                  'BOUNDARY_TABLE':boundary_table,'FLUX_TABLE':flux_table,
                  'MAX_RESIDUAL':f"{max(r['relative_residual'] for r in baseline+strong):.6e}"}.items():
    report = report.replace(key,value)
(Path(__file__).parent/'M1_VALIDATION_REPORT.md').write_text(report,encoding='utf-8')
print(json.dumps({'metadata':metadata,'baseline_finest':[r for r in baseline if r['level']==3],
                  'high_contrast_finest':[r for r in strong if r['level']==3]},indent=2))
