"""M7/M8 reproducible numerical validation and report generation, COMSOL root."""
from pathlib import Path
import sys,json,time
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
ROOT=Path(__file__).resolve().parent
for n in [7,8]:sys.path.insert(0,str(ROOT/f'waveoptics_m{n}'/'src'))
from fiber_modes import Fiber,lp01_analytic,radial_lp01_aeff,solve_modes,scalar_metrics,dispersion_lp01,silica_index
from mcf_hcf import solve_two_core,capillary_neff,capillary_dispersion,run as run_mcf
checks=[]
def chk(name,ok):
 checks.append(dict(name=name,passed=bool(ok)))
 if not ok:raise AssertionError(name)
def stage7():
 p=ROOT/'waveoptics_m7';out=p/'screenshots';out.mkdir(parents=True,exist_ok=True)
 f=Fiber();ref=lp01_analytic(f);oracle=radial_lp01_aeff(f)
 meshes=[]
 for n in (41,61,81):
  sol=solve_modes(f,n);m=scalar_metrics(sol)
  meshes.append(dict(grid=n,triangles=len(sol['triangles']),neff=m['n_eff'],Aeff_um2=m['Aeff_um2'],MFD_um=m['MFD_2nd_um'],error=abs(m['n_eff']-ref['n_eff']),residual=m['residual']))
  chk('M7 guided grid '+str(n),sol['guided'][0]);chk('M7 residual grid '+str(n),m['residual']<1e-8)
 err=[r['error'] for r in meshes];m=meshes[-1]
 chk('M7 independent LP01 neff <1e-5',err[-1]<1e-5)
 chk('M7 mesh refinement trend',err[-1]<err[0] and err[-1]<err[1])
 chk('M7 Aeff versus Bessel oracle within 3%',abs(m['Aeff_um2']/oracle['Aeff_um2']-1)<.03)
 chk('M7 MFD versus Bessel second moment within 3%',abs(m['MFD_um']/oracle['MFD_2nd_um']-1)<.03)
 chk('M7 single-mode V',f.v<2.4048255577)
 chk('M7 analytic dispersion step convergence',abs(dispersion_lp01(step=.003)-dispersion_lp01(step=.0015))<.1)
 chk('M7 silica Malitson',1.443<silica_index(1.55)<1.446)
 wl=np.linspace(1.34,1.66,81);ne=np.array([lp01_analytic(Fiber(wavelength_um=float(w)))['n_eff'] for w in wl])
 d=np.array([dispersion_lp01(w) for w in wl])
 np.savetxt(out/'m7_dispersion.csv',np.column_stack((wl,ne,d)),delimiter=',comments='',header='wavelength_um,LP01_neff,D_ps_nm_km')
 n=sol['n'];x=sol['xy'][:,0].reshape(n,n);y=sol['xy'][:,1].reshape(n,n);I=sol['fields'][:,0].reshape(n,n)**2
 fig,ax=plt.subplots(1,2,figsize=(11,4),constrained_layout=True)
 im=ax[0].pcolormesh(x,y,I,cmap='inferno',shading='auto');ax[0].add_patch(plt.Circle((0,0),f.radius_um,fill=False,edgecolor='cyan'))
 ax[0].set(xlabel='x (um)',ylabel='y (um)',title='M7 scalar FEM LP01',aspect='equal');fig.colorbar(im,ax=ax[0])
 ax[1].plot([r['triangles'] for r in meshes],np.array(err)*1e6,'o-')
 ax[1].set(xlabel='triangles',ylabel='abs neff error x 1e6',title='FEM vs analytical Bessel')
 fig.savefig(out/'m7_mode_and_convergence.png',dpi=145);plt.close(fig)
 fig,ax=plt.subplots(1,2,figsize=(11,4),constrained_layout=True)
 ax[0].plot(wl,ne,label='Bessel LP01');ax[0].plot([1.55],[m['neff']],'o',label='FEM')
 ax[0].set(xlabel='wavelength um',ylabel='neff',title='Spectral guided index');ax[0].legend()
 ax[1].plot(wl,d);ax[1].set(xlabel='wavelength um',ylabel='D ps/(nm km)',title='Analytical LP01+material dispersion')
 fig.savefig(out/'m7_dispersion.png',dpi=145);plt.close(fig)
 data=dict(V=f.v,n_core=f.core,n_clad=f.clad,LP01_analytic_neff=ref['n_eff'],Aeff_analytic_um2=oracle['Aeff_um2'],
           MFD_analytic_um=oracle['MFD_2nd_um'],D_analytic_ps_per_nm_km=dispersion_lp01(),meshes=meshes)
 (out/'m7_summary.json').write_text(json.dumps(data,indent=2))
 lines=['# M7 — 2D P1 FEM Step-Index Optical Fiber','',
        'Weak-guidance scalar eigenproblem: (k0^2 M_n2 - K) psi = beta^2 M psi, PEC-like Dirichlet outer boundary.',
        '1.55um, core radius 4.1um, delta n=0.005, silica Sellmeier. Independent LP01 Bessel analytic oracle.',
        'Aeff=(integral I)^2/integral I^2, MFD=2 sqrt(2<r^2>) by second-moment convention.',
        'Dispersion uses derivative of ANALYTIC LP01 neff, not noisy wavelength-differenced FEM output.',
        '',f'Analytic LP01 neff {ref["n_eff"]:.9f}, Aeff {oracle["Aeff_um2"]:.3f} um2, MFD {oracle["MFD_2nd_um"]:.3f} um.',
        f'Analytic dispersion at 1550nm {data["D_analytic_ps_per_nm_km"]:.3f} ps/(nm km).','',
        '| Grid | triangles | FEM neff | abs error | Aeff um2 | MFD um |','|---:|---:|---:|---:|---:|---:|']
 for r in meshes:lines.append(f'|{r["grid"]}|{r["triangles"]}|{r["neff"]:.9f}|{r["error"]:.3g}|{r["Aeff_um2"]:.3f}|{r["MFD_um"]:.3f}|')
 lines+=['','![M7 FEM field](screenshots/m7_mode_and_convergence.png)','![M7 spectrum](screenshots/m7_dispersion.png)',
         '','LIMITATION: weak-guidance scalar FEM, no vector Maxwell eigenmode, no PML/loss, and no actual manufacturer optical fiber comparison.']
 (p/'M7_VALIDATION_REPORT.md').write_text('\n'.join(lines)+'\n')
 return data
def stage8():
 p=ROOT/'waveoptics_m8';out=p/'screenshots';data=run_mcf(out);m=data['MCF']
 split=[r['index_split'] for r in m]
 chk('M8 positive two-core supermode splitting',all(s>0 for s in split))
 chk('M8 coupling decreases with separation',all(split[i]>split[i+1] for i in range(len(split)-1)))
 chk('M8 even/odd mode parity',all(r['parity'][0]['even_score']<.04 and r['parity'][1]['odd_score']<.04 for r in m))
 chk('M8 supermode residual',all(max(r['residuals'])<1e-6 for r in m))
 conv=[solve_two_core(16,n=n) for n in [81,101,121]]
 data['convergence_16um']=[dict(grid=r['n_grid'],split=r['index_split'],length_mm=r['coupling_length_mm']) for r in conv]
 chk('M8 MCF grid convergence within 7%',abs(conv[-1]['index_split']/conv[0]['index_split']-1)<.07)
 cap=capillary_neff([1.3,1.55,1.7])
 chk('M8 hollow capillary index below air',np.all(cap<1.00027))
 chk('M8 hollow capillary monotonic wavelength',np.all(np.diff(cap)<0))
 chk('M8 capillary large radius limit',abs(float(capillary_neff(1.55,radius_um=1000))-1.00027)<1e-6)
 chk('M8 capillary dispersion finite',np.isfinite(capillary_dispersion()))
 (out/'m8_summary.json').write_text(json.dumps(data,indent=2))
 lines=['# M8 — MCF Coupled Modes and HCF Capillary Baseline','',
        'Two-core symmetric index-guiding weak-guidance scalar P1 FEM; 1.55um, radius 4.1um, delta n=0.005.',
        'Power-transfer length Lc=lambda/[2*(neff_even-neff_odd)].','','| pitch um | even neff | odd neff | delta neff | Lc mm |','|---:|---:|---:|---:|---:|']
 for r in m:lines.append(f'|{r["separation_um"]:.0f}|{r["n_even"]:.9f}|{r["n_odd"]:.9f}|{r["index_split"]:.7g}|{r["coupling_length_mm"]:.3f}|')
 lines+=['','## HCF analytical baseline — NOT NANF/ARF FEM',
         'Ideal hollow capillary: neff=sqrt(n_air^2-(u01*lambda/(2*pi*R))^2), R=15um, n_air=1.00027.',
         f'At 1.55um neff={data["HCF_capillary"]["neff_1p55"]:.9f}; D={data["HCF_capillary"]["D_1p55_ps_per_nm_km"]:.3f} ps/(nm km).',
         'No PML, no imaginary neff, no confinement loss, no nested tubes, no ARF antiresonance.',
         '','![Even/odd supermodes](screenshots/m8_mcf_supermodes.png)',
         '![MCF coupling and HCF](screenshots/m8_coupling_and_hcf.png)',
         '','## MCF convergence at pitch 16um','| grid | splitting | coupling length mm |','|---:|---:|---:|']
 for r in data['convergence_16um']:lines.append(f'|{r["grid"]}|{r["split"]:.7g}|{r["length_mm"]:.3f}|')
 (p/'M8_VALIDATION_REPORT.md').write_text('\n'.join(lines)+'\n')
 return data
if __name__=='__main__':
 t=time.time();a=stage7();b=stage8()
 rec=dict(passed=sum(x['passed'] for x in checks),failed=sum(not x['passed'] for x in checks),checks=checks,M7=a,M8=b,seconds=round(time.time()-t,2))
 (ROOT/'M7_M8_EXECUTION_SUMMARY.json').write_text(json.dumps(rec,indent=2))
 (ROOT/'M1_M8_FINAL_REPORT.md').write_text('# M1–M8 integration status\n\n'
 'M1-M3 archived earlier; M4 partial 3D PEC cavity only; M5/M6 1D optics; M7 validated weak-guidance 2D scalar FEM LP01; M8 validated MCF two-core scalar FEM and HCF capillary analytic baseline.\n\n'
 f'M7/M8 newly passed checks: {rec["passed"]}/{len(checks)}. No COMSOL direct or real fiber specification validation.\n'
 'Full vector Maxwell/PML/complex loss and HCF NANF antiresonant confinement are NOT implemented.\n')
 print(json.dumps(dict(passed=rec['passed'],failed=rec['failed'],neff=a['meshes'][-1]['neff'],Aeff=a['meshes'][-1]['Aeff_um2'],M8_split_16=b['MCF'][1]['index_split']),indent=2))
