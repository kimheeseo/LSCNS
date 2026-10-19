"""Isolated M1-M8 regressions; each milestone has a separate module namespace.

Results are written even on failures. Exit nonzero if any command fails.
Archived full M4 is optional because its 3D meshes cost more CPU and memory.
"""
from pathlib import Path
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import time
from zipfile import ZipFile
ROOT=Path(__file__).resolve().parents[1]


def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--include-archived-m4',action='store_true')
    ap.add_argument('--out',type=Path,default=Path(__file__).resolve().parent/'results')
    args=ap.parse_args();args.out=args.out.resolve();logs=args.out/'regression';logs.mkdir(parents=True,exist_ok=True)
    record=dict(created_utc=datetime.now(timezone.utc).isoformat(),python=sys.version,
                note='Milestone suites include overlapping regressions; do not sum as unique tests.',runs=[])
    env=os.environ.copy();env.update(OPENBLAS_NUM_THREADS='1',OMP_NUM_THREADS='1',MPLBACKEND='Agg')
    def execute(name,cmd,paths=(),cwd=ROOT.parent):
        child=env.copy();child['PYTHONPATH']=os.pathsep.join(map(str,paths))
        t=time.perf_counter()
        with (logs/(name+'.log')).open('w') as stream:
            try:
                result=subprocess.run(cmd,cwd=cwd,env=child,stdout=stream,stderr=subprocess.STDOUT,timeout=900)
                code=result.returncode
            except subprocess.TimeoutExpired:
                stream.write('\nTIMEOUT after 900 seconds\n');code=124
        log=(logs/(name+'.log')).read_text()
        row=dict(name=name,command=cmd,returncode=code,seconds=round(time.perf_counter()-t,2),
                 log_sha256=hashlib.sha256(log.encode()).hexdigest(),log='regression/'+name+'.log',
                 summary=log.splitlines()[-1] if log.splitlines() else '')
        record['runs'].append(row)
        (args.out/'regression_summary.json').write_text(json.dumps(record,indent=2)+'\n')
        print(f'{name}: exit {code}; {row["summary"]}',flush=True)
    for n in (1,2,3):
        p=ROOT/f'waveoptics_m{n}'
        execute(f'M{n}',[sys.executable,'-m','pytest',str(p/'tests'),'-q','--tb=short',f'--junitxml={logs}/M{n}.xml'],[p])
    execute('research',[sys.executable,'-m','pytest',str(ROOT/'research/tests'),'-q','--tb=short',f'--junitxml={logs}/research.xml'])
    execute('M4_M6',[sys.executable,str(ROOT/'run_m4_m6_validation.py')])
    execute('M7_M8',[sys.executable,str(ROOT/'run_m7_m8_validation.py')])
    if args.include_archived_m4:
        archive=ROOT/'reports/M4_3D_2026-10-11/M4_Reproducibility_Source.zip'
        record['archived_m4_sha256']=hashlib.sha256(archive.read_bytes()).hexdigest()
        with tempfile.TemporaryDirectory(prefix='waveoptics-m4-') as tmp:
            with ZipFile(archive) as z:
                for info in z.infolist():
                    rel=Path(info.filename)
                    # Extract only known Python sources and never traverse outside temp.
                    if rel.parts and rel.parts[0]=='waveoptics_m4' and '..' not in rel.parts and rel.suffix=='.py':
                        target=Path(tmp)/rel;target.parent.mkdir(parents=True,exist_ok=True)
                        target.write_bytes(z.read(info.filename))
            p=Path(tmp)/'waveoptics_m4'
            execute('M4_full',[sys.executable,'-m','pytest',str(p/'tests/test_maxwell3d_m4.py'),str(p/'tests/test_m4_extended.py'),
                    '-q','--tb=short',f'--junitxml={logs}/M4_full.xml'],[p,p/'tests'])
    record['all_commands_passed']=all(r['returncode']==0 for r in record['runs'])
    (args.out/'regression_summary.json').write_text(json.dumps(record,indent=2)+'\n')
    return 0 if record['all_commands_passed'] else 1

if __name__=='__main__':sys.exit(main())
