"""Execute the standalone Colab report and preserve real outputs and result ZIP."""
import os
os.environ['OPENBLAS_NUM_THREADS']='1'
os.environ['OMP_NUM_THREADS']='1'
from pathlib import Path
import json,zipfile,argparse,base64
import nbformat
from nbclient import NotebookClient
ROOT=Path(__file__).resolve().parent
if not (ROOT/'G657A1_Colab_Validation.ipynb').exists():
    import build_notebook
nb=nbformat.read(ROOT/'G657A1_Colab_Validation.ipynb',as_version=4)
ap=argparse.ArgumentParser()
ap.add_argument('--in-process',action='store_true',help='Execute cells in IPython when local kernel sockets are unavailable')
args=ap.parse_args()
if args.in_process:
    from IPython.terminal.interactiveshell import TerminalInteractiveShell
    from IPython.utils.capture import capture_output
    shell=TerminalInteractiveShell.instance();os.chdir(ROOT)
    counter=0
    for cell in nb.cells:
        if cell.cell_type!='code':continue
        counter+=1
        existing=set(ROOT.glob('G657A1_Colab_Results/**/*.png'))
        with capture_output() as cap:
            result=shell.run_cell(cell.source,store_history=False)
        if result.error_before_exec or result.error_in_exec:
            print(cap.stdout);print(cap.stderr)
            raise RuntimeError(f'Cell {counter} failed') from (result.error_before_exec or result.error_in_exec)
        cell.execution_count=counter;cell.outputs=[]
        if cap.stdout:cell.outputs.append(nbformat.v4.new_output('stream',name='stdout',text=cap.stdout))
        if cap.stderr:cell.outputs.append(nbformat.v4.new_output('stream',name='stderr',text=cap.stderr))
        for rich in cap.outputs:
            cell.outputs.append(nbformat.v4.new_output('display_data',data=rich.data,metadata=rich.metadata))
        if not any('image/png' in o.get('data',{}) for o in cell.outputs):
            for image_path in sorted(set(ROOT.glob('G657A1_Colab_Results/**/*.png'))-existing):
                cell.outputs.append(nbformat.v4.new_output('display_data',data={'image/png':base64.b64encode(image_path.read_bytes()).decode()},metadata={}))
        print(f'Executed code cell {counter}',flush=True)
    nb.metadata['execution_method']='Actual in-process IPython replay; local kernel socket startup unavailable. Colab UI not manually tested.'
else:
    NotebookClient(nb,timeout=300,kernel_name='python3',resources={'metadata':{'path':str(ROOT)}}).execute()
    nb.metadata['execution_method']='Jupyter kernel via nbclient'

nbformat.write(nb,ROOT/'G657A1_Colab_Validation.ipynb')
with zipfile.ZipFile(ROOT/'G657A1_Colab_Results.zip') as z:
    dest=ROOT/'results';dest.mkdir(exist_ok=True)
    for info in z.infolist():
        rel=Path(info.filename)
        if len(rel.parts)!=2 or rel.parts[0]!='G657A1_Colab_Results':raise ValueError('unexpected archive entry')
        (dest/rel.name).write_bytes(z.read(info.filename))
record=json.loads((ROOT/'results/summary.json').read_text())
print(json.dumps({k:record[k] for k in ('numerical_passed','numerical_total','reference_passed','reference_total','parameters_and_results')},indent=2,ensure_ascii=False))
if record['numerical_passed']!=record['numerical_total']:
    print('NOTICE: failed numerical criteria retained; this report must not be labeled all passed.')
