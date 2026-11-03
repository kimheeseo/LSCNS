"""Check saved examples and replay the editable mesh/wavelength form."""
from pathlib import Path
import hashlib,json,os,subprocess,sys,tempfile
import nbformat
ROOT=Path(__file__).resolve().parent
nb=nbformat.read(ROOT/'Fiber_Input_Output_Examples.ipynb',4);nbformat.validate(nb)
assert not [o for c in nb.cells for o in c.get('outputs',[]) if o.output_type=='error']
assert all(c.execution_count is not None for c in nb.cells if c.cell_type=='code')
assert nb.metadata.engine_sha256==hashlib.sha256((ROOT/'fiber_model.py').read_bytes()).hexdigest()
assert nb.metadata.protocol_sha256==hashlib.sha256((ROOT/'simulation.py').read_bytes()).hexdigest()
baseline=json.loads((ROOT/'results/summary.json').read_text())
with tempfile.TemporaryDirectory(prefix='fiber-io-editable-') as tmp:
    d=Path(tmp)
    for c in nb.cells:
        if c.cell_type=='code' and c.source.startswith('# @title 4.'):
            c.source=c.source.replace('g657a1_wavelength_nm = 1310.0','g657a1_wavelength_nm = 1550.0').replace('"40,80,160,320"','"40,80"').replace('"61,81,121,161"','"41,61"')
    nbformat.write(nb,d/'Fiber_Input_Output_Examples.ipynb')
    (d/'run_report.py').write_bytes((ROOT/'run_report.py').read_bytes())
    child=subprocess.run([sys.executable,str(d/'run_report.py'),'--in-process'],capture_output=True,text=True,timeout=180,env={**os.environ,'OPENBLAS_NUM_THREADS':'1','OMP_NUM_THREADS':'1'})
    (ROOT/'results/edited_parameters_replay.log').write_text(child.stdout+child.stderr)
    assert child.returncode==0,child.stdout+child.stderr
    edited=json.loads((d/'results/summary.json').read_text())
    assert edited['mesh_inputs']['radial_core_counts']==[40,80]
    assert sorted({r['grid_n'] for r in edited['cartesian_mesh']})==[41,61]
    assert sorted({r['core_elements'] for r in edited['radial_mesh']})==[40,80]
    old=next(r for r in baseline['current_outputs'] if r['fiber']=='G.657.A1')
    new=next(r for r in edited['current_outputs'] if r['fiber']=='G.657.A1')
    assert new['analysis_wavelength_nm']==1550 and new['MFD_reference_wavelength_nm']==1310
    assert abs(old['delta_n']-new['delta_n'])<1e-12
    assert abs(new['MFD_PII_at_reference_um']-9.2)<1e-8
    assert abs(new['MFD_PII_um']-9.2)>.5
    # Example E is direct input, not an MFD fit. All example calculations run in both replays.
    manual=next(x for x in edited['examples'] if x['example'].startswith('E:'))
    assert manual['delta_n_source']=='MANUAL' and manual['MFD_target_um'] is None
    assert abs(manual['delta_n']-.0046)<1e-15
    result=dict(notebook_valid=True,all_code_cells_executed=True,engine_and_protocol_hashes_match=True,
       edited_mesh_form_applied_to_actual_solvers=True,analysis_wavelength_1550_keeps_reference_1310=True,
       manual_delta_n_mode_verified=True,fixed_examples_count=len(baseline['examples']),
       actual_edited_grids=[41,61],actual_edited_radial_core_counts=[40,80],
       colab_ui_manually_tested=False,execution_method='Actual in-process IPython cell replay')
    (ROOT/'results/report_validation.json').write_text(json.dumps(result,indent=2)+'\n')
print(json.dumps(result,indent=2))
