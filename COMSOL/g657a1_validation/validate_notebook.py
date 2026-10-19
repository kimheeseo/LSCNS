"""Validate real notebook outputs and the independent 1310-nm calibration anchor."""
from pathlib import Path
import hashlib,json,os,subprocess,sys,tempfile,zipfile
import nbformat
ROOT=Path(__file__).resolve().parent
nb=nbformat.read(ROOT/'G657A1_Colab_Validation.ipynb',4)
nbformat.validate(nb)
code=[c for c in nb.cells if c.cell_type=='code']
assert all(c.execution_count is not None for c in code)
assert not [o for c in code for o in c.outputs if o.output_type=='error']
images=sum('image/png' in o.get('data',{}) for c in code for o in c.outputs)
assert images==8,images
engine=(ROOT/'fiber_model.py').read_text()
assert nb.metadata.source_engine_sha256==hashlib.sha256(engine.encode()).hexdigest()
assert any(c.source.endswith(engine) for c in code)
base=json.loads((ROOT/'results/summary.json').read_text())
# A known 2D discretization failure must remain visible, not be turned into green.
assert base['numerical_passed']==sum(base['numerical_checks'].values())
assert base['numerical_total']==len(base['numerical_checks'])
assert not base['numerical_checks']['2D |neff error| <1e-5']
with tempfile.TemporaryDirectory(prefix='g657a1-anchor-') as tmp:
    d=Path(tmp)
    for cell in nb.cells:
        if cell.cell_type=='code' and cell.source.startswith('# @title 3.'):
            cell.source=cell.source.replace('wavelength_nm = 1310.0','wavelength_nm = 1550.0',1)
            cell.source=cell.source.replace('run_2d_fem = True','run_2d_fem = False').replace('run_profile_sensitivity = True','run_profile_sensitivity = False')
    nbformat.write(nb,d/'G657A1_Colab_Validation.ipynb')
    (d/'g657a1_validation.py').write_bytes((ROOT/'g657a1_validation.py').read_bytes())
    child=subprocess.run([sys.executable,str(d/'g657a1_validation.py'),'--in-process'],capture_output=True,text=True,timeout=180,env={**os.environ,'OPENBLAS_NUM_THREADS':'1','OMP_NUM_THREADS':'1'})
    (ROOT/'results/parameter_replay.log').write_text(child.stdout+child.stderr)
    assert child.returncode==0,child.stderr
    alt=json.loads((d/'results/summary.json').read_text())
    a=alt['parameters_and_results'];b=base['parameters_and_results']
    assert abs(a['fitted_delta_n']-b['fitted_delta_n'])<1e-12
    assert abs(a['MFD_PII_1310_um']-9.2)<1e-8
    assert abs(a['Aeff_um2']-b['Aeff_1550_um2'])<1e-8
    assert abs(a['D1550_ps_nm_km']-b['D1550_ps_nm_km'])<1e-6
    assert a['analysis_wavelength_nm']==1550 and a['MFD_reference_wavelength_nm']==1310
    assert alt['numerical_total']==8
    with zipfile.ZipFile(d/'G657A1_Colab_Results.zip') as z:
        pngs=[x for x in z.namelist() if x.endswith('.png')]
        assert len(pngs)==5,len(pngs) # no stale 2D/sensitivity images from the saved default notebook
    result=dict(notebook_schema_valid=True,executed_code_cells=len(code),embedded_pngs=images,
       no_error_outputs=True,engine_sha256_verified=True,known_2d_failure_preserved=True,
       analysis_1550_keeps_mfd_anchor_1310=True,delta_n_difference=a['fitted_delta_n']-b['fitted_delta_n'],
       alternate_numerical_passed=alt['numerical_passed'],alternate_numerical_total=alt['numerical_total'],
       alternate_export_pngs=5,colab_ui_manually_tested=False,
       tested_execution_method='Actual in-process IPython replay; local socket kernel unavailable')
    (ROOT/'results/notebook_validation.json').write_text(json.dumps(result,indent=2)+'\n')
print(json.dumps(result,indent=2))
