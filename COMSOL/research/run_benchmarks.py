"""Run research qualification without rewriting historical milestone evidence."""
from pathlib import Path
import argparse
from dataclasses import replace
from datetime import datetime, timezone
import hashlib
import json
import os
import platform
import subprocess
import sys
import time
import numpy as np
import scipy
ROOT=Path(__file__).resolve().parents[1]
for stage in (7,8):sys.path.insert(0,str(ROOT/f'waveoptics_m{stage}'/'src'))
from fiber_modes import Fiber,lp01_analytic,radial_lp01_aeff,solve_modes,scalar_metrics,dispersion_lp01
from radial_fem import solve_radial,radial_dispersion
from mcf_hcf import solve_two_core


def run(out):
    out.mkdir(parents=True,exist_ok=True)
    started=time.perf_counter()
    record=dict(schema_version=1,created_utc=datetime.now(timezone.utc).isoformat(),
                model='weak-guidance scalar step-index; constant delta_n',
                python=sys.version,platform=platform.platform(),numpy=np.__version__,scipy=scipy.__version__,
                threads={k:os.environ.get(k) for k in ('OPENBLAS_NUM_THREADS','OMP_NUM_THREADS')})
    record['source_sha256']={str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest()
        for p in sorted(ROOT.rglob('*.py')) if '__pycache__' not in p.parts}
    record['radial_cases']=[]
    cases=[(1.31,3.5,.004),(1.55,4.1,.005),(1.625,5.,.003),(1.55,5.5,.002)]
    for w,a,dn in cases:
        f=Fiber(wavelength_um=w,radius_um=a,delta_n=dn,domain_halfwidth_um=50.)
        ref=dict(**lp01_analytic(f),**radial_lp01_aeff(f))
        mesh=[]
        for count in (40,80,160,320):
            s=solve_radial(f,core_elements=count,clad_elements=4*count)
            keep={k:v for k,v in s.items() if k not in ('radius_um','field')}
            keep.update(neff_error=abs(s['n_eff']-ref['n_eff']),
                        Aeff_relative_error=abs(s['Aeff_um2']/ref['Aeff_um2']-1),
                        MFD_relative_error=abs(s['MFD_2nd_um']/ref['MFD_2nd_um']-1))
            mesh.append(keep)
        last=mesh[-1]
        passed=last['neff_error']<2e-7 and last['Aeff_relative_error']<5e-4 and last['MFD_relative_error']<5e-4 and last['residual']<1e-9
        record['radial_cases'].append(dict(wavelength_um=w,radius_um=a,delta_n=dn,V=f.v,reference=ref,meshes=mesh,passed=passed))
    f=Fiber()
    record['cartesian_mesh']=[dict(grid=n,**scalar_metrics(solve_modes(f,n=n,modes=1))) for n in (41,61,81,121,161)]
    # Fixed 0.3 um spacing separates outer-box changes from mesh changes.
    record['cartesian_domain']=[dict(halfwidth_um=h,grid=n,**scalar_metrics(solve_modes(replace(f,domain_halfwidth_um=h),n=n,modes=1))) for h,n in ((18.,121),(24.,161),(30.,201))]
    record['radial_domain']=[{k:v for k,v in solve_radial(f,core_elements=160,clad_elements=int((R-f.radius_um)/.05),outer_radius_um=R).items() if k not in ('radius_um','field')} for R in (18.,26.,40.)]
    record['dispersion']=[dict(core_elements=n,clad_elements=4*n,**radial_dispersion(replace(f,domain_halfwidth_um=40.),h,n,4*n)) for n in (80,160,320) for h in (.01,.005,.0025)]
    record['dispersion_analytic_ps_nm_km']=dispersion_lp01()
    record['mcf']=[]
    for pitch in (12.,16.,20.,24.):
        entries=[]
        for n in (81,121,161):
            s=solve_two_core(pitch,n=n)
            entries.append({k:v for k,v in s.items() if k!='solution'})
        # Same fine spacing, enlarged box: 54/160 = 72/(214-1) is approximate;
        # choose 216 intervals for h=1/3 against 162 intervals below.
        a=solve_two_core(pitch,n=163,halfwidth_um=27.)
        b=solve_two_core(pitch,n=217,halfwidth_um=36.)
        mesh_change=abs(entries[-1]['index_split']/entries[-2]['index_split']-1)
        box_change=abs(b['index_split']/a['index_split']-1)
        record['mcf'].append(dict(pitch_um=pitch,meshes=entries,
          domain=[{k:v for k,v in s.items() if k!='solution'} for s in (a,b)],
          relative_mesh_change=mesh_change,relative_domain_change=box_change,
          screening_passed=mesh_change<.05 and box_change<.05,
          qualification='5 percent screening only; not a rigorous uncertainty bound or stochastic XT prediction'))
    record['elapsed_seconds']=time.perf_counter()-started
    (out/'benchmarks.json').write_text(json.dumps(record,indent=2,allow_nan=False)+'\n')
    lines=['# Research qualification benchmarks','',f"Created: {record['created_utc']}",'',
           'Raw inputs, meshes, numerical values, environment and source hashes: [benchmarks.json](benchmarks.json).',
           'These are scalar-model verification results. No direct COMSOL run or manufacturer measurement comparison.',
           '', '| wavelength um | core radius um | delta n | V | finest neff error | Aeff rel error | MFD rel error | gate |',
           '|---:|---:|---:|---:|---:|---:|---:|---|']
    for c in record['radial_cases']:
        s=c['meshes'][-1]
        lines.append(f"|{c['wavelength_um']}|{c['radius_um']}|{c['delta_n']}|{c['V']:.4f}|{s['neff_error']:.3g}|{s['Aeff_relative_error']:.3g}|{s['MFD_relative_error']:.3g}|{'PASS' if c['passed'] else 'FAIL'}|")
    lines += ['', 'Radial acceptance: |delta neff| <2e-7, relative Aeff and second-moment MFD errors <5e-4; residual <1e-9. Thresholds fixed in runner.',
              '', '## Two-core MCF numerical screening', '', '| Pitch um | last mesh change | fixed-spacing box change | screening |','|---:|---:|---:|---|']
    for c in record['mcf']:
        lines.append(f"|{c['pitch_um']}|{c['relative_mesh_change']:.3%}|{c['relative_domain_change']:.3%}|{'PASS' if c['screening_passed'] else 'FAIL'}|")
    lines += ['', 'This screening uses a 5% threshold for BOTH mesh and box changes. It is not an error bound. Failure excludes that configuration from quantitative coupling claims.',
              'Lc=lambda/(2 delta neff) is the ideal identical-core complete power-transfer length. It is not deployed-cable crosstalk.',
              '', '## Dispersion', '', '| core elements | wavelength step um | D ps/(nm km) | difference from Bessel |','|---:|---:|---:|---:|']
    for c in record['dispersion']:
        lines.append(f"|{c['core_elements']}|{c['step_um']}|{c['D_ps_nm_km']:.8f}|{c['D_ps_nm_km']-record['dispersion_analytic_ps_nm_km']:.6g}|")
    (out/'BENCHMARK_REPORT.md').write_text('\n'.join(lines)+'\n')
    return record

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--out',type=Path,default=Path(__file__).resolve().parent/'results')
    args=p.parse_args();r=run(args.out)
    print(json.dumps(dict(radial_passed=sum(c['passed'] for c in r['radial_cases']),radial_cases=len(r['radial_cases']),
                         MCF_screening=[{'pitch':c['pitch_um'],'passed':c['screening_passed']} for c in r['mcf']],seconds=r['elapsed_seconds']),indent=2))
    # MCF screening failures are reported, not silently converted into success claims.
    if not all(c['passed'] for c in r['radial_cases']):sys.exit(1)
