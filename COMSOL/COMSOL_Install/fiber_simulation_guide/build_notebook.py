"""Build a self-contained, editable input/output and mesh demonstration notebook."""
from pathlib import Path
import hashlib
import nbformat
ROOT=Path(__file__).resolve().parent
nb=nbformat.v4.new_notebook();cells=[]
def md(t):cells.append(nbformat.v4.new_markdown_cell(t))
def code(t):cells.append(nbformat.v4.new_code_cell(t))
md('''# 광섬유 시뮬레이션 — 입력 → 출력 → 예시 → Mesh 검증

**G.654.E / G.657.A1 통합 실행형 보고서 · 2026-10-19**

[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/fiber_simulation_guide/Fiber_Input_Output_Examples.ipynb)

`런타임 → 모두 실행`으로 두 광섬유의 기본 예시, 입력 변경 예시, radial/2D Mesh 비교와 외부 영역 검사를 실제 재계산합니다. **기본 예시의 코드 실행 출력과 그림도 저장**했습니다.

이 도구는 가정한 원형 step-index 단면의 약유도 스칼라 LP01 계산기입니다. 실제 제조사 굴절률 구조 또는 COMSOL 결과의 독립 검증은 아닙니다.
''')
md('''## 1. 어떤 값을 입력하는가?

| 입력 | 단위·의미 | 사용 방법 |
|---|---|---|
| 분석 파장 | nm | 해당 파장에서 모드·MFD·Aeff·D를 계산 |
| 코어 반경 a | µm, 지름 아님 | 기본값 G.654.E 5.5, G.657.A1 4.1: **가정값** |
| Δn | n_core−n_clad, 절대 차이 | 직접 계산 모드에서 사용 |
| 목표 MFD | µm, Petermann-II | 보정 모드에서만 입력; Δn을 맞춤 |
| MFD 기준 파장 | G.654.E 1550 / G.657.A1 1310 nm | 분석 파장을 바꿔도 기준은 유지 |
| 클래딩 굴절률 | 실리카 Sellmeier로 자동 계산 | n_core=n_clad+일정 Δn 가정 |
| n2 | m²/W | gamma용 **가정값**, 직접 변경 가능 |
| Radial 요소 수 / 클래딩 비율 | 코어 요소 수와 클래딩 요소 수의 비 | Mesh 세분화 검사 |
| 2D N×N / 외곽 반폭 | 홀수 N / µm | 계산영역의 실제 전체 폭은 반폭의 두 배 |
| Radial 외곽 반경 | µm | Mesh 수와 별도로 경계 영향 검사 |

**직접 계산:** a·Δn 입력 → MFD가 출력됩니다. **MFD 보정:** a·목표 MFD 입력 → Δn이 보정됩니다. 보정한 MFD를 독립적으로 예측한 값이라고 보고하지 않습니다.
''')
md('''## 2. 어떤 결과가 나오는가?

| 출력 | 의미·단위 | 주의 |
|---|---|---|
| n_core / n_clad / V / n_eff | 굴절률, 정규화 주파수, 유효굴절률 | LP01 기준 모델 |
| 2D 모드 세기 | 상대 | Bessel 이론 그림과 실제 FEM 그림을 구분 |
| MFD_PII / MFD_2nd | µm | Petermann-II와 중심 2차 모멘트 정의 구분 |
| Aeff | µm² | 장분포에서 직접 적분; MFD 원 면적과 일반적으로 다름 |
| D / S | ps/(nm·km) / ps/(nm²·km) | 이론 n_eff의 파장 미분, 실제 도핑 분산 미사용 |
| λ0 / S0 | nm / ps/(nm²·km) | 지원 파장 구간에서 영점을 찾은 경우만 표시; 미해결은 빈 값 |
| gamma | 1/(W·km) | 입력 n2와 예측 Aeff를 사용한 파생값 |
| FEM 오차 / 잔차 | 절대오차·상대오차(%) | 같은 가상 단면의 Bessel 해와 비교 |
| Mesh·영역 표 | 설정과 계산값 | Mesh를 늘렸다는 사실만으로 정확도 보장 불가 |

**계산하지 않는 값:** 실제 감쇠, 굽힘 손실, 케이블 컷오프 인증, PMD·복굴절, 전체 전송거리·BER·GN 전송성능. 별도 물리 모델이나 측정이 필요합니다.
''')
code('''# @title 3. 환경 설정
import os,sys,json,uuid,zipfile,platform
from pathlib import Path
import numpy as np
import pandas as pd
import scipy
import matplotlib.pyplot as plt
from IPython.display import display,FileLink
get_ipython().run_line_magic('matplotlib','inline')
RESULT_DIR=Path('Fiber_Input_Output_Results')/uuid.uuid4().hex[:12]
RESULT_DIR.mkdir(parents=True,exist_ok=False)
print('Python',sys.version.split()[0],'NumPy',np.__version__,'SciPy',scipy.__version__)
''')
code('''# @title 4. 입력값 — 이 셀을 바꾼 뒤 아래 셀들을 다시 실행
fit_delta_n_from_mfd = True # @param {type:"boolean"}
g654e_wavelength_nm = 1550.0 # @param {type:"number"}
g654e_core_radius_um = 5.5 # @param {type:"number"}
g654e_target_mfd_um = 12.5 # @param {type:"number"}
g654e_manual_delta_n = 0.0034785793846155 # @param {type:"number"}
g654e_n2_m2_W = 2.2e-20 # @param {type:"number"}
g657a1_wavelength_nm = 1310.0 # @param {type:"number"}
g657a1_core_radius_um = 4.1 # @param {type:"number"}
g657a1_target_mfd_um = 9.2 # @param {type:"number"}
g657a1_manual_delta_n = 0.0046215295453259 # @param {type:"number"}
g657a1_n2_m2_W = 2.6e-20 # @param {type:"number"}
radial_core_mesh_text = "40,80,160,320" # @param {type:"string"}
cladding_to_core_element_ratio = 4 # @param {type:"integer"}
radial_outer_radius_um = 40.0 # @param {type:"number"}
mesh_2d_text = "61,81,121,161" # @param {type:"string"}
box_2d_halfwidth_um = 24.0 # @param {type:"number"}
run_2d_fem = True # @param {type:"boolean"}
# MFD基準波長は製品資料に対応する固定値です。
MFD_REFERENCE_NM={'G.654.E':1550.,'G.657.A1':1310.}
'''.replace('# MFD基準波長は製品資料に対応する固定値です。','# MFD 기준 파장은 공개 자료의 값에 대응하도록 고정합니다.'))
engine=(ROOT/'fiber_model.py').read_text()
helpers=(ROOT/'simulation.py').read_text().replace('from fiber_model import Fiber, silica_index, lp01_analytic, radial_lp01_aeff, solve_radial, solve_modes, scalar_metrics, dispersion_lp01\n','')
code('# @title 5. 자체 포함 FEM 엔진 (펼쳐보기)\n'+engine+'\n'+helpers)
code('''# @title 6. 현재 입력 → 현재 출력
configs={
 'G.654.E':dict(analysis_wavelength_nm=g654e_wavelength_nm,reference_wavelength_nm=1550.,radius_um=g654e_core_radius_um,target_mfd_um=g654e_target_mfd_um,manual_delta_n=g654e_manual_delta_n,n2_m2_W=g654e_n2_m2_W),
 'G.657.A1':dict(analysis_wavelength_nm=g657a1_wavelength_nm,reference_wavelength_nm=1310.,radius_um=g657a1_core_radius_um,target_mfd_um=g657a1_target_mfd_um,manual_delta_n=g657a1_manual_delta_n,n2_m2_W=g657a1_n2_m2_W)}
input_df=pd.DataFrame(configs).T;display(input_df)
fibers={};current_rows=[]
for name,cfg in configs.items():
    f,row=summarize(cfg,fit_delta_n_from_mfd,radial_outer_radius_um)
    fibers[name]=f;current_rows.append(dict(fiber=name,**row))
current_df=pd.DataFrame(current_rows)
display(current_df.T)
print('MFD_TARGET는 보정 입력, MFD_PII_at_reference는 보정 결과입니다. 같은 값이라는 사실은 제품 독립 검증이 아닙니다.')
''')
md('''## 7. 기본 예시와 입력 변경 예시

아래 예시는 현재 설정과 별개로 고정된 입력을 사용합니다. (1) 두 기본 단면, (2) 분석 파장 변경, (3) 목표 MFD를 사용하지 않는 Δn 직접 입력, (4) 코어 반경 가정 변경을 비교합니다. MFD 보정의 기준 파장은 그대로 유지합니다.
''')
code('''# @title 7-1. 예시를 실제로 계산
examples=[]
case_specs=[
 ('A: G654E baseline','G.654.E',{},True),
 ('B: G657A1 baseline','G.657.A1',{},True),
 ('C: G654E analysis 1610 nm','G.654.E',{'analysis_wavelength_nm':1610.},True),
 ('D: G657A1 analysis 1550 nm','G.657.A1',{'analysis_wavelength_nm':1550.},True),
 ('E: G657A1 manual delta_n','G.657.A1',{'analysis_wavelength_nm':1550.,'manual_delta_n':.0046},False),
 ('F: G657A1 assumed radius 3.8 um','G.657.A1',{'analysis_wavelength_nm':1550.,'radius_um':3.8},True)]
for label,name,changes,fit in case_specs:
    cfg={**PRESETS[name],**changes};_,row=summarize(cfg,fit,40.)
    examples.append(dict(example=label,fiber=name,**row))
examples_df=pd.DataFrame(examples)
display(examples_df[['example','analysis_wavelength_nm','MFD_reference_wavelength_nm','assumed_radius_um','delta_n','delta_n_source','MFD_PII_at_reference_um','MFD_PII_um','Aeff_um2','D_ps_nm_km','gamma_W_inv_km_inv']])
# Semantic checks: changing analysis wavelength must not re-fit the catalog at a new wavelength.
assert abs(examples[0]['delta_n']-examples[2]['delta_n'])<1e-12
assert abs(examples[1]['delta_n']-examples[3]['delta_n'])<1e-12
assert examples[4]['delta_n_source']=='MANUAL' and examples[4]['MFD_target_um'] is None
assert abs(examples[1]['MFD_PII_at_reference_um']-examples[5]['MFD_PII_at_reference_um'])<1e-8
assert abs(examples[3]['Aeff_um2']-examples[5]['Aeff_um2'])>1
''')
md('''## 8. Mesh 수에 따른 변화 — 계산영역 고정

Radial 외곽 반경과 2D 외곽 반폭을 각각 고정합니다. 두 솔버의 요소 수는 서로 같은 의미가 아닙니다. Radial은 원형 대칭 문제이고, 2D는 삼각형으로 단면을 풉니다. **한 Mesh의 가장 작은 오차만 골라 솔버 전체 정확도로 보고하면 안 됩니다.**
''')
code('''# @title 8-1. 두 단면의 실제 FEM 메시 재계산
core_counts=parse_counts(radial_core_mesh_text)
grids=parse_counts(mesh_2d_text,odd=True)
radial_frames=[];cart_frames=[];last_2d={};checks=[]
fig,axs=plt.subplots(2,2,figsize=(11,7.2))
for i,(name,f) in enumerate(fibers.items()):
    rad,cart,last=mesh_study(f,core_counts,int(cladding_to_core_element_ratio),grids,box_2d_halfwidth_um,run_2d_fem)
    radial_frames.append(rad.assign(fiber=name));last_2d[name]=last
    display(rad.assign(fiber=name))
    axs[i,0].semilogy(rad.total_elements,rad.neff_abs_error,'o-');axs[i,0].set(title=name+' radial: fixed domain',xlabel='Total radial elements',ylabel='Absolute neff error')
    rr=rad.iloc[-1]
    checks.extend([dict(fiber=name,test='radial |neff error| <1e-6',passed=bool(rr.neff_abs_error<1e-6)),dict(fiber=name,test='radial |Aeff error| <1%',passed=bool(abs(rr.Aeff_error_pct)<1))])
    if len(cart):
        cart_frames.append(cart.assign(fiber=name));display(cart.assign(fiber=name))
        axs[i,1].semilogy(cart.grid_n,cart.neff_abs_error,'o-');axs[i,1].axhline(1e-5,color='gray',ls='--')
        axs[i,1].set(title=name+' 2D centroid interface',xlabel='Grid N x N',ylabel='Absolute neff error')
        checks.append(dict(fiber=name,test='2D |neff error| <1e-5',passed=bool(cart.iloc[-1].neff_abs_error<1e-5)))
    else:axs[i,1].text(.1,.5,'2D FEM disabled',transform=axs[i,1].transAxes)
for ax in axs.flat:ax.grid(alpha=.25)
fig.tight_layout();fig.savefig(RESULT_DIR/'01_Mesh_Comparison.png',dpi=140);plt.show()
radial_df=pd.concat(radial_frames,ignore_index=True)
cartesian_df=pd.concat(cart_frames,ignore_index=True) if cart_frames else pd.DataFrame()
''')
code('''# @title 8-2. 실제 2D FEM 세기 분포 (2D 계산을 켠 경우)
if run_2d_fem:
    fig,axs=plt.subplots(1,2,figsize=(10,4.5))
    for ax,(name,s) in zip(axs,last_2d.items()):
        n=s['n'];xy=s['xy'];I=s['fields'][:,0].reshape(n,n)**2
        im=ax.pcolormesh(xy[:,0].reshape(n,n),xy[:,1].reshape(n,n),I/I.max(),shading='auto',cmap='inferno')
        ax.add_patch(plt.Circle((0,0),s['fiber'].radius_um,fill=False,edgecolor='white',ls='--'))
        ax.set(title=name+f' actual 2D FEM: {n}x{n}',xlabel='x (um)',ylabel='y (um)',aspect='equal');fig.colorbar(im,ax=ax)
    fig.tight_layout();fig.savefig(RESULT_DIR/'02_Actual_2D_Fields.png',dpi=140);plt.show()
''')
md('''## 9. 計算영역 크기 검사 — Mesh 간격 유지

외부 경계가 너무 가까우면 장의 꼬리를 잘라 오차를 만듭니다. 코어 요소 수를 고정하고 클래딩 간격을 약 0.04 µm로 유지하면서 외곽 반경을 바꿉니다. 이것은 Mesh 세분화 검사와 별개입니다.
'''.replace('計算','계산'))
code('''# @title 9-1. Radial 외곽 반경 30/40/50 um 비교
domain_frames=[]
for name,f in fibers.items():
    d=domain_study(f,core_elements=max(core_counts))
    domain_frames.append(d.assign(fiber=name));display(d.assign(fiber=name))
    checks.append(dict(fiber=name,test='domain last neff change <1e-7',passed=bool(abs(d.iloc[-1].neff-d.iloc[-2].neff)<1e-7)))
domain_df=pd.concat(domain_frames,ignore_index=True)
''')
code('''# @title 10. 파장별 MFD / Aeff / D 예측
spectra=[];fig,axs=plt.subplots(1,3,figsize=(13,3.8))
for name,f in fibers.items():
    lo,hi=(1530,1625) if name=='G.654.E' else (1260,1625)
    rows=[]
    for wl in np.linspace(lo,hi,45):
        ff=replace(f,wavelength_um=wl/1000);mm=radial_lp01_aeff(ff)
        rows.append(dict(fiber=name,wavelength_nm=float(wl),MFD_PII_um=mm['MFD_petermann_ii_um'],Aeff_um2=mm['Aeff_um2'],D_ps_nm_km=dispersion_lp01(ff)))
    d=pd.DataFrame(rows);spectra.append(d)
    for ax,y in zip(axs,['MFD_PII_um','Aeff_um2','D_ps_nm_km']):ax.plot(d.wavelength_nm,d[y],label=name);ax.set(xlabel='Wavelength (nm)',ylabel=y);ax.grid(alpha=.25);ax.legend(fontsize=8)
fig.tight_layout();fig.savefig(RESULT_DIR/'03_Spectral_Outputs.png',dpi=140);plt.show()
spectrum_df=pd.concat(spectra,ignore_index=True)
''')
md('''## 11. 결과의 의미와 사용 제한

- Δn 직접 입력 모드: MFD는 모델 출력입니다. MFD 보정 모드: 목표 MFD와 일치하는 것은 보정의 결과입니다.
- 코어 반경·Δn은 제품 이름만으로 결정되지 않습니다. 같은 MFD를 만족하는 여러 가상 단면이 나옵니다.
- 원형 경계 정합 FEM과 기존 2D 중심점 분류 방식은 기하 오차가 다릅니다. 2D 오차가 Mesh에 대해 비단조적일 수 있습니다.
- 스칼라 근사, 파장에 대해 일정한 Δn, 유한 계산영역의 영향이 있습니다. 실제 제품·굽힘 손실·전체 벡터 해석 정확도를 이 결과만으로 주장하지 않습니다.
- MFD_PII와 MFD_2nd는 정의가 다릅니다. Aeff를 단순히 pi*(MFD/2)^2로 바꾸지 않습니다.

원본 제품별 실행 보고서: [G.654.E](https://github.com/kimheeseo/LSCNS/tree/main/COMSOL/g654e_validation) · [G.657.A1](https://github.com/kimheeseo/LSCNS/tree/main/COMSOL/g657a1_validation)
''')
code('''# @title 12. 입력·출력·예시·검증 결과 저장 및 ZIP
checks_df=pd.DataFrame(checks);display(checks_df)
print('이 수치 판정은 같은 가상 단면의 해석해 대비 결과입니다. FAIL은 그대로 저장합니다.')
frames={'current_inputs.csv':input_df.reset_index(names='fiber'),'current_outputs.csv':current_df,'example_inputs_outputs.csv':examples_df,
 'radial_mesh.csv':radial_df,'radial_domain.csv':domain_df,'spectral_outputs.csv':spectrum_df,'checks.csv':checks_df}
if len(cartesian_df):frames['cartesian_mesh.csv']=cartesian_df
for filename,df in frames.items():df.to_csv(RESULT_DIR/filename,index=False)
def records(df):return json.loads(df.to_json(orient='records',double_precision=15))
record=dict(schema_version=1,model='ideal circular step-index weak-guidance scalar LP01',
 MFD_fit_enabled=fit_delta_n_from_mfd,current_inputs=configs,current_outputs=records(current_df),examples=records(examples_df),
 numerical_checks=records(checks_df),passed=int(checks_df.passed.sum()),total=len(checks_df),
 radial_mesh=records(radial_df),cartesian_mesh=records(cartesian_df),radial_domain=records(domain_df),
 mesh_inputs=dict(radial_core_counts=core_counts,clad_ratio=cladding_to_core_element_ratio,radial_outer_radius_um=radial_outer_radius_um,grid_counts=grids,box_halfwidth_um=box_2d_halfwidth_um,run_2d=run_2d_fem),
 environment=dict(python=sys.version,numpy=np.__version__,scipy=scipy.__version__,pandas=pd.__version__,matplotlib=plt.matplotlib.__version__,platform=platform.platform()),
 not_computed=['actual attenuation','bending loss','measured cable cutoff','PMD/birefringence','BER/reach','full-vector Maxwell equivalence'])
(RESULT_DIR/'summary.json').write_text(json.dumps(record,ensure_ascii=False,indent=2,allow_nan=False),encoding='utf-8')
report='# 입력/출력 실제 실행 보고서\\n\\n'
for x in record['current_outputs']:
    report+=f"- {x['fiber']}: analysis={x['analysis_wavelength_nm']} nm, a={x['assumed_radius_um']} um, Delta n={x['delta_n']:.10f}, Aeff={x['Aeff_um2']:.5f} um2, D={x['D_ps_nm_km']:.5f} ps/(nm km)\\n"
report+=f"\\nNumerical checks: {record['passed']}/{record['total']} PASS; failures preserved. Calibration is not independent product validation.\\n"
(RESULT_DIR/'EXECUTION_REPORT.md').write_text(report,encoding='utf-8')
archive=Path('Fiber_Input_Output_Results.zip')
with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as z:
    for file in sorted(RESULT_DIR.iterdir()):
        if file.is_file():z.write(file,arcname='Fiber_Input_Output_Results/'+file.name)
display(FileLink(str(archive)));print('현재 실행에서 만든 파일만 ZIP에 포함:',archive)
# from google.colab import files
# files.download(str(archive))
''')
nb.cells=cells
nb.metadata={'kernelspec':{'display_name':'Python 3','language':'python','name':'python3'},'language_info':{'name':'python'},'colab':{'name':'Fiber_Input_Output_Examples.ipynb','provenance':[]},'engine_sha256':hashlib.sha256(engine.encode()).hexdigest(),'protocol_sha256':hashlib.sha256((ROOT/'simulation.py').read_bytes()).hexdigest()}
nbformat.write(nb,ROOT/'Fiber_Input_Output_Examples.ipynb')
print('Built',len(cells),'cells')
