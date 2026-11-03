"""Execute real self-contained notebook cells without a TCP Jupyter kernel.
Preserve stream, display Markdown and PNG outputs in the notebook.
"""
from pathlib import Path
import io,contextlib,hashlib,json,time
import nbformat
ROOT=Path(__file__).resolve().parents[1]


def replay(path=None,edited=False):
    path=Path(path) if path else ROOT/'M8_HCF_MCF_2Core_Validation.ipynb'
    nb=nbformat.read(path,as_version=4);ns={'__name__':'__notebook__'}
    if edited:
        nb.cells[3].source=nb.cells[3].source.replace('pitch_um = 16.0','pitch_um = 18.0').replace('hcf_radius_um = 15.0','hcf_radius_um = 20.0')
    from IPython.display import display
    import IPython.display as ipd
    original=ipd.display;start=time.time();count=0
    for cell in nb.cells:
        if cell.cell_type!='code':continue
        count+=1;cell.outputs=[];cell.execution_count=count
        def capture(*objects,**kwargs):
            for obj in objects:
                if hasattr(obj,'_repr_png_'):
                    data=obj._repr_png_()
                    if isinstance(data,tuple):data=data[0]
                    import base64
                    cell.outputs.append(nbformat.v4.new_output('display_data',data={'image/png':data if isinstance(data,str) else base64.b64encode(data).decode(),'text/plain':'<PNG image>'}))
                elif hasattr(obj,'_repr_markdown_'):
                    cell.outputs.append(nbformat.v4.new_output('display_data',data={'text/markdown':obj._repr_markdown_(),'text/plain':str(obj)}))
                else:cell.outputs.append(nbformat.v4.new_output('display_data',data={'text/plain':str(obj)}))
        ipd.display=capture
        stream=io.StringIO()
        try:
            with contextlib.redirect_stdout(stream):exec(compile(cell.source,f'<notebook-cell-{count}>','exec'),ns)
        finally:ipd.display=original
        if stream.getvalue():cell.outputs.insert(0,nbformat.v4.new_output('stream',name='stdout',text=stream.getvalue()))
        print('Executed cell',count,flush=True)
    result=ns['result'];assert result['passed']==result['total'],result['checks']
    if edited:
        assert result['config']['pitch_um']==18. and result['config']['hcf_radius_um']==20.
        base=json.loads((ROOT/'validation'/'results'/'summary.json').read_text())
        assert abs(result['mcf_finest']['coupling_length_mm']-base['mcf_finest']['coupling_length_mm'])>1
        assert abs(result['hcf_finest']['D_ps_nm_km']-base['hcf_finest']['D_ps_nm_km'])>1
    else:
        nbformat.validate(nb);nbformat.write(nb,path)
        import shutil
        shutil.copy2(ns['PACKAGE'],ROOT/'M8_HCF_MCF_Results.zip')
    report=dict(edited_parameters=edited,executed_code_cells=count,checks_passed=result['passed'],checks_total=result['total'],
                png_outputs=sum('image/png' in o.get('data',{}) for c in nb.cells for o in c.get('outputs',[])),
                pitch_um=result['config']['pitch_um'],hcf_radius_um=result['config']['hcf_radius_um'],
                coupling_length_mm=result['mcf_finest']['coupling_length_mm'],hcf_D=result['hcf_finest']['D_ps_nm_km'],
                runtime_seconds=time.time()-start,colab_browser_tested=False)
    if not edited:
        report['embedded_hashes_match']=all(hashlib.sha256((ROOT.parent/key).read_bytes()).hexdigest()==value for key,value in nb.metadata.embedded_source_sha256.items())
        assert report['embedded_hashes_match']
    (ROOT/'validation'/'results'/('edited_notebook_replay.json' if edited else 'notebook_replay.json')).write_text(json.dumps(report,indent=2)+'\n')
    print(json.dumps(report,indent=2));return report


if __name__=='__main__':
    import argparse
    p=argparse.ArgumentParser();p.add_argument('--edited',action='store_true');args=p.parse_args();replay(edited=args.edited)
