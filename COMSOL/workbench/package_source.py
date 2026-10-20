"""Build a relocatable source bundle with the exact M7/M8 dependencies."""
from pathlib import Path
import zipfile,hashlib,json
ROOT=Path(__file__).resolve().parent;COMSOL=ROOT.parent
DEPENDENCIES=['waveoptics_m7/src/fiber_modes.py','waveoptics_m7/src/radial_fem.py','waveoptics_m8/src/capillary_fem.py','waveoptics_m8/src/conforming_mcf.py','waveoptics_m8/src/multipole_reference.py']


def build():
    files=[p for p in sorted(ROOT.rglob('*')) if p.is_file() and not set(p.parts)&{'__pycache__','.venv','.pytest_cache'} and p.suffix not in ['.zip','.pyc'] and p.name!='SOURCE_MANIFEST.json']
    files += [COMSOL/p for p in DEPENDENCIES]
    manifest=dict(version='0.1.0',date_kst='2026-10-20',root='COMSOL',files={str(p.relative_to(COMSOL)):{'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in files})
    (ROOT/'SOURCE_MANIFEST.json').write_text(json.dumps(manifest,indent=2,ensure_ascii=False)+'\n')
    files.append(ROOT/'SOURCE_MANIFEST.json')
    package=ROOT/'LS_WaveOptics_Workbench_v0.1.zip'
    prefix='LS_WaveOptics_Workbench_v0.1/'
    with zipfile.ZipFile(package,'w',zipfile.ZIP_DEFLATED) as z:
        for p in files:z.write(p,prefix+'COMSOL/'+str(p.relative_to(COMSOL)))
        z.writestr(prefix+'START_HERE.txt','LS Wave Optics Workbench v0.1\n\n1. Install Python 3.12.\n2. On Windows, open COMSOL/workbench/START_WINDOWS.bat.\n3. Open http://127.0.0.1:8765/ and confirm the Python engine is connected.\n\nFirst setup needs internet for public Python packages. Computation then runs locally.\nRead COMSOL/workbench/README.md for model assumptions and validation limits.\n')
    with zipfile.ZipFile(package) as z:assert z.testzip() is None
    print(json.dumps(dict(package=str(package),bytes=package.stat().st_size,source_files=len(files))))
    return package


if __name__=='__main__':build()
