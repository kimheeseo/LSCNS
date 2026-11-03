"""Reproducible HCF baseline and independent two-core MCF validation.
Run from anywhere: python COMSOL/waveoptics_m8/validation/run_validation.py
"""
from pathlib import Path
import sys,json,csv,time,platform
ROOT=Path(__file__).resolve().parents[1]
sys.path[:0]=[str(ROOT/'src'),str(ROOT.parent/'waveoptics_m7'/'src')]
import numpy as np
from scipy.special import j0,jn_zeros
from scipy.integrate import quad
from fiber_modes import Fiber,lp01_analytic,C0
from mcf_hcf import capillary_neff,capillary_dispersion,solve_two_core
from capillary_fem import solve_capillary
from conforming_mcf import solve_conforming,coherent_transfer
from multipole_reference import solve_reference

DEFAULT_CONFIG=dict(wavelength_um=1.55,core_radius_um=4.1,delta_n=.005,pitch_um=16.,
                    mesh_sizes_um=[1.,.7,.45,.3],halfwidth_um=40.,domain_halfwidths_um=[28.,32.,40.],
                    pitches_um=[12.,16.,20.,24.],hcf_radius_um=15.,n_air=1.00027,
                    hcf_elements=[40,80,160,320,640])


def write_csv(path,rows):
    with Path(path).open('w',newline='',encoding='utf-8') as f:
        w=csv.DictWriter(f,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)


def relative_percent(value,reference):return abs(value-reference)/abs(reference)*100


def run_validation(output_dir,config=None,include_legacy=True):
    start=time.time();out=Path(output_dir);out.mkdir(parents=True,exist_ok=True)
    cfg={**DEFAULT_CONFIG,**(config or {})}
    f=Fiber(wavelength_um=cfg['wavelength_um'],radius_um=cfg['core_radius_um'],delta_n=cfg['delta_n'])
    pitch=cfg['pitch_um'];wl=f.wavelength_um
    checks=[]
    def check(name,value,limit):
        checks.append(dict(name=name,value=float(value),limit=float(limit),passed=bool(np.isfinite(value) and value<=limit)))
    # HCF: the FEM eigenvalue is independent of analytic Bessel zeros.
    radius=cfg['hcf_radius_um'];gas=cfg['n_air'];u=float(jn_zeros(0,1)[0]);kt2=(u/radius)**2
    exact_n=float(capillary_neff(wl,radius,gas));exact_D=float(capillary_dispersion(wl,radius,gas))
    p=quad(lambda t:j0(u*t)**2*t,0,1,epsabs=1e-13)[0]
    p4=quad(lambda t:j0(u*t)**4*t,0,1,epsabs=1e-13)[0]
    exact_A=2*np.pi*radius**2*p*p/p4
    hcf=[]
    for count in cfg['hcf_elements']:
        row=solve_capillary(wl,radius,gas,count)
        row.update(neff_abs_error=abs(row['n_eff']-exact_n),neff_error_pct=relative_percent(row['n_eff'],exact_n),
                   kt2_error_pct=relative_percent(row['transverse_kt2_per_um2'],kt2),
                   D_error_pct=relative_percent(row['D_ps_nm_km'],exact_D),Aeff_error_pct=relative_percent(row['Aeff_um2'],exact_A))
        hcf.append(row)
    finest=hcf[-1]
    check('HCF neff absolute error',finest['neff_abs_error'],1e-9)
    check('HCF transverse eigenvalue relative error (%)',finest['kt2_error_pct'],.0001)
    check('HCF D relative error (%)',finest['D_error_pct'],.0001)
    check('HCF Aeff relative error (%)',finest['Aeff_error_pct'],.0001)
    check('HCF eigen residual',finest['residual'],1e-8)
    check('HCF refinement neff error ratio',finest['neff_abs_error']/hcf[0]['neff_abs_error'],.05)
    # Separate derivative check: five independent FEM neff computations.
    derivative=[]
    for h in [.02,.01,.005]:
        vals=[solve_capillary(wl+j*h,radius,gas,cfg['hcf_elements'][-1])['n_eff'] for j in [-2,-1,0,1,2]]
        d2=np.dot([-1,16,-30,16,-1],vals)/(12*h*h)
        D=-wl/C0*1e12*d2
        derivative.append(dict(step_um=h,D_FEM_fivepoint=float(D),error_vs_analytic_pct=relative_percent(D,exact_D),
                               difference_vs_discrete_exact_pct=relative_percent(D,finest['D_ps_nm_km'])))
    check('HCF finite-difference D relative error (%)',max(x['error_vs_analytic_pct'] for x in derivative),.001)
    # Wider radius/wavelength verification, without parameter fitting.
    sweep=[]
    for a in [10.,15.,30.]:
        for w in [1.3,1.55,1.7]:
            q=solve_capillary(w,a,gas,cfg['hcf_elements'][-1]);n=float(capillary_neff(w,a,gas));d=float(capillary_dispersion(w,a,gas))
            sweep.append(dict(radius_um=a,wavelength_um=w,neff_FEM=q['n_eff'],neff_reference=n,
                              neff_abs_error=abs(q['n_eff']-n),D_FEM=q['D_ps_nm_km'],D_reference=d,D_error_pct=relative_percent(q['D_ps_nm_km'],d)))
    check('HCF 9-case max neff absolute error',max(x['neff_abs_error'] for x in sweep),1e-9)
    check('HCF 9-case max D relative error (%)',max(x['D_error_pct'] for x in sweep),.0001)
    # Independent infinite-cladding reference. Truncation and angular sampling tested.
    refs=[solve_reference(f,pitch,order=m) for m in [2,4,6,8]]
    ref=refs[-1];sample=solve_reference(f,pitch,order=8,samples=512)
    check('MCF multipole 6->8 splitting difference (%)',relative_percent(ref['index_split'],refs[-2]['index_split']),.00001)
    check('MCF angular 256->512 splitting difference (%)',relative_percent(sample['index_split'],ref['index_split']),.00001)
    far=solve_reference(f,60.,order=6);single=lp01_analytic(f)['n_eff']
    check('MCF large-pitch single-core neff limit',max(abs(far[k]-single) for k in ('n_even','n_odd')),1e-9)
    def with_errors(row,reference):
        row=dict(row)
        row.update(neff_even_abs_error=abs(row['n_even']-reference['n_even']),neff_odd_abs_error=abs(row['n_odd']-reference['n_odd']),
                   split_error_pct=relative_percent(row['index_split'],reference['index_split']),
                   Lc_error_pct=relative_percent(row['coupling_length_mm'],reference['coupling_length_mm']))
        return row
    cache={}
    def solve(d,L,h,field=False):
        key=(d,L,h)
        if key not in cache or field:
            cache[key]=solve_conforming(f,d,L,h,return_fields=field)
        return cache[key]
    meshes=[]
    for h in cfg['mesh_sizes_um']:
        row=with_errors(solve(pitch,cfg['halfwidth_um'],h),ref);meshes.append(row)
        print(f'MCF mesh h={h}: Lc={row["coupling_length_mm"]:.9g} mm, error={row["Lc_error_pct"]:.6g}%',flush=True)
    final=meshes[-1]
    check('MCF finest max neff absolute error',max(final['neff_even_abs_error'],final['neff_odd_abs_error']),1e-7)
    check('MCF finest splitting relative error (%)',final['split_error_pct'],.05)
    check('MCF finest Lc relative error (%)',final['Lc_error_pct'],.05)
    check('MCF finest residual',max(final['residuals']),1e-9)
    check('MCF refinement Lc error ratio',final['Lc_error_pct']/meshes[0]['Lc_error_pct'],.1)
    domains=[]
    # Same nominal h for all domains; independent remesh, not identical topology.
    h=cfg['mesh_sizes_um'][-1]
    for L in cfg['domain_halfwidths_um']:
        domains.append(with_errors(solve(pitch,L,h),ref))
    check('MCF last two domain splitting change (%)',relative_percent(domains[-1]['index_split'],domains[-2]['index_split']),.01)
    pitches=[]
    for d in cfg['pitches_um']:
        r=solve_reference(f,d,order=8)
        z=with_errors(solve(d,cfg['halfwidth_um'],h),r)
        z.update(reference_split=r['index_split'],reference_Lc_mm=r['coupling_length_mm']);pitches.append(z)
    check('MCF pitch sweep max Lc error (%)',max(z['Lc_error_pct'] for z in pitches),.05)
    check('MCF pitch sweep coupling monotonic violations',sum(pitches[i]['index_split']<=pitches[i+1]['index_split'] for i in range(len(pitches)-1)),0)
    legacy=[]
    if include_legacy:
        for n in [81,101,121]:
            z=solve_two_core(pitch,n,27.,fiber=f)
            z.pop('solution');legacy.append(with_errors(z,ref))
    length=np.linspace(0,2*final['coupling_length_mm']*1e-3,401)
    transfer=coherent_transfer(length,final['coupling_length_mm'])
    check('MCF coherent power conservation',np.max(np.abs(transfer['P1']+transfer['P2']-1)),1e-14)
    special=coherent_transfer(np.array([0,1,2])*final['coupling_length_mm']*1e-3,final['coupling_length_mm'])['P2']
    check('MCF coherent transfer at 0,Lc,2Lc',np.max(abs(special-[0,1,0])),1e-12)
    # Plot actual P2 FEM fields, not an illustrative Gaussian.
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    fig,ax=plt.subplots(1,3,figsize=(13,3.8),constrained_layout=True)
    ax[0].loglog([z['elements'] for z in hcf],[z['neff_abs_error'] for z in hcf],'o-');ax[0].set(xlabel='Radial elements',ylabel='Absolute n_eff error',title='HCF ideal capillary: P1 convergence')
    ax[1].loglog([z['quarter_dofs'] for z in meshes],[z['Lc_error_pct'] for z in meshes],'o-');ax[1].set(xlabel='Quarter-domain P2 DOFs',ylabel='Coupling-length error (%)',title='MCF vs infinite-domain reference')
    ax[2].semilogy([z['halfwidth_um'] for z in domains],[z['Lc_error_pct'] for z in domains],'o-');ax[2].set(xlabel='Domain halfwidth (um)',ylabel='Coupling-length error (%)',title='MCF domain study (remeshed)')
    for a in ax:a.grid(alpha=.3)
    fig.savefig(out/'01_convergence.png',dpi=150);plt.close(fig)
    sol=solve(pitch,cfg['halfwidth_um'],cfg['mesh_sizes_um'][-1],field=True)
    # Plot the actual P2 nodal values on four subtriangles per element.
    # The curved solver mesh is retained; linear plotting is only a display step.
    ed=sol['basis'].element_dofs
    triangles=np.concatenate([ed[[0,3,5]].T,ed[[3,1,4]].T,ed[[5,4,2]].T,ed[[3,4,5]].T])
    points=sol['basis'].doflocs
    centers=points[:,triangles].mean(axis=2)
    triangles=triangles[(centers[0]<pitch/2+12)&(centers[1]<13)]
    fig,ax=plt.subplots(1,2,figsize=(11,4),constrained_layout=True)
    for i,parity in enumerate([1,-1]):
        q=sol['fields'][i]/np.max(abs(sol['fields'][i]))
        for sx,sy in [(1,1),(-1,1),(1,-1),(-1,-1)]:
            im=ax[i].tripcolor(sx*points[0],sy*points[1],triangles,q*(parity if sx<0 else 1),shading='gouraud',cmap='RdBu_r',vmin=-1,vmax=1,rasterized=True)
        for c in [-pitch/2,pitch/2]:ax[i].add_patch(plt.Circle((c,0),f.radius_um,fill=False,color='k',lw=.7))
        ax[i].set(xlabel='x (um)',ylabel='y (um)',title=['Even P2 FEM supermode','Odd P2 FEM supermode'][i],aspect='equal',xlim=(-pitch/2-12,pitch/2+12),ylim=(-13,13));fig.colorbar(im,ax=ax[i],label='Normalized scalar field')
    fig.savefig(out/'02_mcf_supermodes.png',dpi=150);plt.close(fig)
    fig,ax=plt.subplots(1,2,figsize=(11,4),constrained_layout=True)
    ax[0].semilogy([z['pitch_um'] for z in pitches],[z['coupling_length_mm'] for z in pitches],'o',label='P2 FEM');ax[0].semilogy([z['pitch_um'] for z in pitches],[z['reference_Lc_mm'] for z in pitches],'-',label='Fourier-Bessel');ax[0].set(xlabel='Core pitch (um)',ylabel='Lc (mm)',title='Identical straight two-core model');ax[0].legend()
    ax[1].plot(length*1000,transfer['P1'],label='Core 1');ax[1].plot(length*1000,transfer['P2'],label='Core 2');ax[1].set(xlabel='Length (mm)',ylabel='Power fraction',title='Coherent transfer: not random cable XT');ax[1].legend()
    for a in ax:a.grid(alpha=.3)
    fig.savefig(out/'03_coupling.png',dpi=150);plt.close(fig)
    for name,rows in [('hcf_mesh',hcf),('hcf_derivative',derivative),('hcf_sweep',sweep),('mcf_reference_orders',refs),('mcf_mesh',meshes),('mcf_domain',domains),('mcf_pitch',pitches),('checks',checks)]:write_csv(out/(name+'.csv'),rows)
    if legacy:write_csv(out/'mcf_legacy.csv',legacy)
    write_csv(out/'mcf_transfer.csv',[dict(length_mm=float(z*1000),P1=float(a),P2=float(b)) for z,a,b in zip(length,transfer['P1'],transfer['P2'])])
    summary=dict(config=cfg,scope='scalar-model numerical verification; no experimental HCF/MCF validation',
                 hcf_reference=dict(n_eff=exact_n,D_ps_nm_km=exact_D,kt2=kt2,Aeff_um2=exact_A),
                 hcf_finest=finest,hcf_mesh=hcf,hcf_derivative=derivative,hcf_sweep=sweep,
                 mcf_reference=ref,mcf_reference_orders=refs,mcf_finest=final,mcf_mesh=meshes,mcf_domain=domains,mcf_pitch=pitches,mcf_legacy=legacy,
                 checks=checks,passed=sum(z['passed'] for z in checks),total=len(checks),runtime_seconds=time.time()-start,
                 python=platform.python_version(),colab_browser_tested=False)
    (out/'summary.json').write_text(json.dumps(summary,indent=2,ensure_ascii=False)+'\n')
    print(f'Checks: {summary["passed"]}/{summary["total"]}; runtime {summary["runtime_seconds"]:.1f}s',flush=True)
    return summary


if __name__=='__main__':
    import argparse
    parser=argparse.ArgumentParser();parser.add_argument('--output',default=str(ROOT/'validation'/'results'));parser.add_argument('--skip-legacy',action='store_true')
    args=parser.parse_args();result=run_validation(args.output,include_legacy=not args.skip_legacy)
    if result['passed']!=result['total']:raise SystemExit('Numerical validation has failed gates; see checks.csv')
