"""Colab 호환 노트북의 셀을 IPython에서 순서대로 실행하고 출력을 저장한다."""
from pathlib import Path
import os, time
import nbformat
from IPython.core.interactiveshell import InteractiveShell
from IPython.utils.capture import capture_output

def main():
    path=Path(__file__).with_name('GN_Model_Colab.ipynb').resolve()
    notebook=nbformat.read(path,as_version=4)
    os.chdir(path.parent)
    shell=InteractiveShell.instance()
    count=0
    for cell in notebook.cells:
        if cell.cell_type!='code': continue
        count+=1;start=time.perf_counter()
        with capture_output() as captured:
            result=shell.run_cell(cell.source,store_history=True)
        if result.error_before_exec or result.error_in_exec:
            raise RuntimeError(str(result.error_before_exec or result.error_in_exec))
        outputs=[]
        if captured.stdout: outputs.append(nbformat.v4.new_output('stream',name='stdout',text=captured.stdout))
        if captured.stderr: outputs.append(nbformat.v4.new_output('stream',name='stderr',text=captured.stderr))
        outputs.extend(nbformat.v4.new_output('display_data',data=x.data,metadata=x.metadata) for x in captured.outputs)
        cell.execution_count=count;cell.outputs=outputs
        cell.metadata['validation_timing']={'elapsed_seconds':round(time.perf_counter()-start,3)}
        nbformat.write(notebook,path)
        print(f'Cell {count}: executed',flush=True)
    notebook.metadata['execution_environment']='Actual sequential cell execution using IPython in-process; not a Google Colab session.'
    nbformat.validate(notebook);nbformat.write(notebook,path)

if __name__=='__main__': main()
