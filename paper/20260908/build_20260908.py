"""Build and actually execute the self-contained canonical 20260908 notebook.

Uses a real in-process IPython kernel because this runtime disallows ZMQ port
binding. No simulated cell outputs and no claim of execution in Colab's UI.
"""
from pathlib import Path
import os, sys, json, hashlib, datetime, importlib.metadata
import nbformat as nbf

ROOT=Path(__file__).resolve().parent
os.chdir(ROOT)

def md(s):return nbf.v4.new_markdown_cell(s.strip())
def code(s):return nbf.v4.new_code_cell(s.strip(),metadata={'trusted':True})

cells=[md('''# 20260908 · HCF2 손실 모델 수정 및 실제 실행 검증

**실측값에 fitting하지 않은 계산**을 비교합니다. 10–15% 목표 달성을 보장하지 않습니다.
수정 주 모델은 **vector 7영역 동심 근사**입니다. 기존 scalar 5영역/7영역과 추가 vector 9영역 형상 가정도 모두 표시합니다.

아래 출력은 제공된 Python/IPython 커널에서 실제 실행되었습니다. **Colab 웹에서 실행한 것은 아닙니다.**
파일을 열면 실행된 표와 그래프를 바로 볼 수 있습니다. Colab에서 모두 재실행할 수 있도록 코드와 원자료 CSV를 내장했습니다.

비교 대상은 Petrovich et al., Nature Photonics 19, 1203–1208 (2025), DOI [10.1038/s41566-025-01747-5](https://doi.org/10.1038/s41566-025-01747-5)의 **HCF2**입니다.
계산값은 **leakage only**, 논문값은 **measured total**입니다. 이 차이를 총손실 예측 정확도로 오해하면 안 됩니다.
'''),code('''import sys, subprocess, importlib.util
required = {'numpy':'numpy', 'scipy':'scipy', 'pandas':'pandas',
            'matplotlib':'matplotlib', 'mpmath':'mpmath'}
missing = [pkg for mod,pkg in required.items() if importlib.util.find_spec(mod) is None]
if missing:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', *missing])
from pathlib import Path
import json, io, hashlib, types, platform, datetime, importlib.metadata
import numpy as np
import pandas as pd
from IPython.display import display, Image, Markdown
OUT=Path('validated_results');OUT.mkdir(exist_ok=True)
environment={'python':sys.version,'platform':platform.platform(),
             'packages':{p:importlib.metadata.version(p) for p in
                         ['numpy','scipy','pandas','matplotlib','mpmath','ipython']},
             'execution':'actual IPython kernel; Colab web UI not used',
             'executed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat()}
display(environment)
'''),md('''## 구현 코드와 입력 동결

Bird Eqs. (4)–(8)의 네 접선장 경계조건과 exact Bessel/Hankel 기저를 사용합니다.
코드에는 HCF2 손실 보정계수가 없습니다. 아래 세 모듈은 저장소 소스를 그대로 내장한 것입니다.
''')]
for name in ['hcf_concentric_tmm_20260908','hcf_vector_tmm','validation_pipeline']:
    source=(ROOT/(name+'.py')).read_text()
    cells.append(code(f"# Embedded source: {name}.py\nMODULE_SOURCE = {source!r}\nmodule = types.ModuleType({name!r})\nmodule.__file__ = {name+'.py'!r}\nsys.modules[{name!r}] = module\nexec(compile(MODULE_SOURCE, module.__file__, 'exec'), module.__dict__)\nprint('Loaded {name}; SHA256:', hashlib.sha256(MODULE_SOURCE.encode()).hexdigest())"))
cells += [code('''from validation_pipeline import *
frozen_inputs={'geometry_um':GEOMETRY,'evaluation_wavelengths_nm':WAVELENGTHS_NM.tolist(),
               'primary_model':PRIMARY_MODEL,'fitted_loss_parameters':[],
               'blind_validation':False,
               'source_commit':'634a69728e824622777be31f01e99f66b5eafa66'}
(OUT/'frozen_inputs.json').write_text(json.dumps(frozen_inputs,indent=2))
display(frozen_inputs)
'''),md('''코어 반경 14.75 µm, 막 두께 약 0.50 µm를 유지합니다. 큰/중간/작은 튜브 지름은 공개 범위의 중간값 31.05/23.75/7.70 µm입니다.
내부 접촉을 가정한 환산 g1=6.30, g2=15.05 µm는 **실측 gap이 아닙니다**.
실제 5개 튜브 배치가 동심 근사에서는 사라집니다. 치수를 손실에 맞춰 역산하지 않았습니다.
'''),code('''# No paper attenuation data has entered a calculation function.
predictions=predict_spectrum()
predictions.to_csv(OUT/'all_model_predictions.csv',index=False)
assert predictions.converged.all()
assert len(predictions)==46*4
display(predictions.groupby('model').agg(n=('wavelength_nm','size'),
                                        min_loss=('loss_db_km','min'),max_loss=('loss_db_km','max')))
'''),md('''## 수치 구현 검증 — HCF2 실측 정확도와 구분

Bird Table 1의 HE11 정확 수치값 네 개를 Eq. (3)의 동일 기하에서 재계산합니다.
Table 2의 손실 최소화 치수는 사용하지 않습니다. 40자리 별도 구현은 동일 Maxwell 문제의 수치 일치만 검사합니다.
'''),code('''bird=bird_benchmark()
bird.to_csv(OUT/'bird_table1_validation.csv',index=False)
assert bird.converged.all() and bird.within_table_rounding.all()
display(bird)
print('Maximum difference from rounded Bird values (%):',bird.relative_error_pct.abs().max())
'''),code('''precision=numerical_checks(predictions)
precision.to_csv(OUT/'arbitrary_precision_checks.csv',index=False)
display(precision)
print('Maximum float64 vs 40-digit difference (%):',precision.difference_pct.abs().max())
assert precision.difference_pct.abs().max()<0.01
'''),code('''single_wall=single_wall_diagnostic()
single_wall.to_csv(OUT/'bache_single_wall_checks.csv',index=False)
display(single_wall)
print('Bache FEM fitting factor applied: NO')
'''),md('''## 논문 원자료 로딩과 오차 계산

Fig. 3 Source Data의 `HCF_cutback_loss`만 읽습니다. `gas_free_HCF_loss`나 Fig. 1 보정용 섬유는 비교하지 않습니다.
1799점 전체를 보존하고, 1200–1650 nm에서 10 nm 간격 **46점**을 보간 없이 비교합니다.
원문의 1310/1550 nm **반올림값**에 대한 2점 평가는 원자료 평가와 별도입니다.

부호 있는 상대오차 = (계산−논문)/논문 × 100; APE = |계산−논문|/|논문| × 100.
논문값이 이전 대화에서 알려져 있었으므로 완전한 blind 검증이라고 부르지 않습니다.
''')]
raw=(ROOT/'data/petrovich_hcf2_cutback_source.csv').read_text()
cells.append(code(f"RAW_SOURCE_CSV = {raw!r}\nraw=pd.read_csv(io.StringIO(RAW_SOURCE_CSV))\nassert len(raw)==1799 and raw.wavelength_nm.is_unique\nassert np.isfinite(raw.to_numpy()).all()\nraw.to_csv(OUT/'petrovich_hcf2_cutback_source.csv',index=False)\nraw_evaluation=raw[raw.wavelength_nm.isin(WAVELENGTHS_NM)].copy()\nassert len(raw_evaluation)==46\ndisplay(raw[raw.wavelength_nm.isin([1310.,1550.])])\nprint('Digitization: none. Exact source-data samples, no interpolation.')"))
cells += [code('''text_reference=pd.DataFrame({'wavelength_nm':[1310.,1550.],
                             'paper_loss_db_km':[.128,.091]})
nominal=compare(predictions,text_reference,'published text rounded; n=2')
spectral=compare(predictions,raw_evaluation,'source-data exact samples; n=46')
nominal_metrics=metrics(nominal);spectral_metrics=metrics(spectral)
nominal.to_csv(OUT/'all_models_nominal_comparison.csv',index=False)
spectral.to_csv(OUT/'all_models_spectral_comparison.csv',index=False)
primary_comparison_table(nominal).to_csv('20260908_nominal.csv',index=False)
primary_comparison_table(spectral).to_csv('20260908_errors.csv',index=False)
display(primary_comparison_table(nominal))
display(nominal_metrics)
display(spectral_metrics)
'''),code('''# Every evaluation point is visible; individual failures cannot hide behind MAPE.
with pd.option_context('display.max_rows',None,'display.max_columns',None):
    display(primary_comparison_table(spectral))
'''),md('''## 기하 불확실성과 모든 모델의 비교

공개 튜브 지름의 하한/중간/상한으로 만든 27개 조합을 모두 계산합니다.
최소오차 조합을 선택하지 않습니다. 이 범위는 신뢰구간이나 실제 단면들의 공동 분포가 아닙니다.
별도 9영역 가정의 결과도 그래프에 표시하며, 잘 맞는 파장에서만 주 모델을 바꾸지 않습니다.
'''),code('''sensitivity=geometry_sensitivity()
assert sensitivity.converged.all() and len(sensitivity)==54
sensitivity.to_csv(OUT/'geometry_sensitivity_all_27_combinations.csv',index=False)
display(sensitivity.groupby('wavelength_nm').loss_db_km.agg(['count','min','median','max']))
'''),code('''figure_paths=make_plots(predictions,raw,spectral,nominal,sensitivity,Path('.'))
for path in figure_paths:
    display(Image(filename=str(path)))
'''),md((ROOT/'SOURCES_AND_METHODS.md').read_text()),code('''primary=nominal_metrics[nominal_metrics.model==PRIMARY_MODEL].iloc[0]
wide=spectral_metrics[spectral_metrics.model==PRIMARY_MODEL].iloc[0]
old=nominal_metrics[nominal_metrics.model=='old_scalar5'].iloc[0]
summary={'primary_model':PRIMARY_MODEL,
         'quantity_warning':'Calculated leakage compared with measured total; not total-loss prediction accuracy',
         'nominal_metrics':nominal_metrics.to_dict(orient='records'),
         'spectral_metrics':spectral_metrics.to_dict(orient='records'),
         'nominal_primary':primary_comparison_table(nominal).to_dict(orient='records'),
         'target_10pct_all_points_met':bool(wide.n_exceeds_10pct==0),
         'target_15pct_all_points_met':bool(wide.n_within_15pct==wide.n),
         'bird_max_APE_pct':float(bird.relative_error_pct.abs().max()),
         'precision_max_difference_pct':float(precision.difference_pct.abs().max()),
         'can_replace_Petrovich_FEM_or_total_loss':False,
         'environment':environment}
(OUT/'summary.json').write_text(json.dumps(summary,indent=2,ensure_ascii=False))
report=f"""수정 주 모델은 vector 7영역 동심 근사입니다. 본문 반올림값 2점 기준 MAPE는 **{primary.MAPE_pct:,.2f}%**, 최대 APE는 **{primary.max_APE_pct:,.2f}%**입니다.

1200–1650 nm 원자료 46점 기준 MAPE는 **{wide.MAPE_pct:,.2f}%**, 최대 APE는 **{wide.max_APE_pct:,.2f}%**, 10% 초과는 **{int(wide.n_exceeds_10pct)}/{int(wide.n)}점**입니다.
기존 scalar 5영역의 같은 2점 MAPE는 **{old.MAPE_pct:,.2f}%**입니다. **10–15% 목표에 도달하지 못했습니다.**

Bird 동심 문제는 표 반올림 범위 안에서 재현했습니다. 이는 DNANF의 실제 5개 튜브, 방위각 구속, 모드 결합 및 종방향 변동을 대체하지 않습니다.
현재 leakage가 측정 total보다 큰 상태에서 양의 산란·미세굽힘 항을 더하는 것으로 과대예측을 해결할 수 없습니다.
실제 SEM 단면을 사용한 full-vector FEM/PML 수렴 검증이 우선이고, 이어서 독립 측정한 표면/굽힘/가스 정보를 더해야 합니다.

실행 위치: 제공된 Python/IPython 커널. Colab 웹 실행을 주장하지 않습니다. 실제 출력이 포함된 20260908.ipynb와 HTML 및 PNG가 제공됩니다."""
display(Markdown(report))
(OUT/'report_ko.md').write_text(report)
''')]

nb=nbf.v4.new_notebook(cells=cells,metadata={
    'kernelspec':{'display_name':'Python 3','language':'python','name':'python3'},
    'language_info':{'name':'python','version':sys.version.split()[0]},
    'colab':{'name':'20260908.ipynb','provenance':[]},
    'execution_method':'actual IPython in-process kernel; not Colab web UI'})
nbpath=ROOT/'20260908.ipynb'
nbf.write(nb,nbpath)

os.environ['PYDEVD_DISABLE_FILE_VALIDATION']='1'
from ipykernel.inprocess.ipkernel import InProcessKernel
from IPython.utils.capture import capture_output
kernel=InProcessKernel()
shell=kernel.shell
execution=[]
for i,cell in enumerate(nb.cells):
    if cell.cell_type!='code':continue
    print(f'Executing notebook cell {i+1}/{len(nb.cells)}',flush=True)
    started=datetime.datetime.now(datetime.timezone.utc)
    count=shell.execution_count
    with capture_output() as captured:
        result=shell.run_cell(cell.source,store_history=True)
    outputs=[]
    if captured.stdout:outputs.append(nbf.v4.new_output('stream',name='stdout',text=captured.stdout))
    if captured.stderr:outputs.append(nbf.v4.new_output('stream',name='stderr',text=captured.stderr))
    for rich in captured.outputs:
        outputs.append(nbf.v4.new_output('display_data',data=rich.data,metadata=rich.metadata))
    cell.outputs=outputs;cell.execution_count=count
    execution.append({'cell_index':i,'execution_count':count,'success':bool(result.success),
                      'elapsed_seconds':(datetime.datetime.now(datetime.timezone.utc)-started).total_seconds()})
    nbf.write(nb,nbpath)  # retain executed checkpoints if a later cell fails
    if not result.success:
        print(captured.stdout,flush=True);print(captured.stderr,flush=True)
        raise RuntimeError(f'Cell {i} failed: {result.error_in_exec or result.error_before_exec}')
    print(f'  complete; {len(outputs)} outputs',flush=True)
nbf.validate(nb)
codes=[c for c in nb.cells if c.cell_type=='code']
assert all(c.execution_count is not None for c in codes)
assert not any(o.output_type=='error' for c in codes for o in c.outputs)
image_count=sum('image/png' in o.get('data',{}) for c in codes for o in c.outputs)
assert image_count>=4
audit={'all_code_cells_executed':True,'code_cells':len(codes),
       'embedded_png_outputs':image_count,'execution':execution,
       'notebook_sha256':hashlib.sha256(nbpath.read_bytes()).hexdigest(),
       'execution_method':'actual IPython in-process kernel; Colab web UI not used'}
(ROOT/'validated_results/execution_audit.json').write_text(json.dumps(audit,indent=2))
from nbconvert import HTMLExporter
exporter=HTMLExporter();exporter.exclude_input=True
html,_=exporter.from_notebook_node(nb)
(ROOT/'20260908.html').write_text(html)
print(json.dumps({k:v for k,v in audit.items() if k!='execution'},indent=2),flush=True)
