"""Executed numerical evidence; independent references live only in tests/."""
from pathlib import Path
from datetime import datetime,timezone
import csv
import hashlib
import json
import platform
import time
import importlib.metadata as metadata
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from tests.test_vector_m2 import rectangle,fiber
from tests.analytic_vector import rectangle_spectrum,FiberReference,relative_field_error
from tests.test_m2_extended import absorbing_fiber_reference,finite_pec_fiber_reference
from waveoptics.vector import fiber_mesh,solve_vector

OUT = Path(__file__).resolve().parent/'results'
OUT.mkdir(exist_ok=True)


def row(sol,m,label,exact,field_error=None):
    return dict(label=label,neff_real=float(m.neff.real),neff_imag=float(m.neff.imag),
                exact_real=float(exact.real),exact_imag=float(exact.imag),
                neff_abs_error=float(abs(m.neff-exact)),area_um2=m.effective_area_um2,
                loss_db_per_m=m.loss_db_per_m,pencil_residual=m.pencil_residual,
                gauss_residual=m.gauss_residual,field_relative_l2=field_error)


def fiber_field_error(sol,mode,ref,exact):
    xy,w,e,_ = sol.quadrature_fields(mode)
    f = np.stack([ref.electric(xy,exact,angle) for angle in (0,np.pi/2)])
    mat = (f*np.sqrt(w)[None,None,...]).reshape(2,-1).T
    target = (e*np.sqrt(w)).ravel()
    coeff = np.linalg.lstsq(mat,target,rcond=None)[0]
    return float(np.linalg.norm(target-mat@coeff)/np.linalg.norm(target))


def slopes(levels,key,errors):
    h = np.array([r[key] for r in levels])
    err = np.asarray(errors)
    return np.log(err[:-1]/err[1:])/np.log(h[:-1]/h[1:])


def field_plot(sol,filename,title,limit=None):
    m = sol.modes[0]
    et = sol.edge_basis.interpolate(m.edge_coefficients)
    ez = -1j*m.beta_per_um*sol.scalar_basis.interpolate(m.scalar_coefficients)
    values = [abs(et[0])**2,abs(et[1])**2,abs(ez)**2,
              np.sum(abs(et)**2,axis=0)+abs(ez)**2]
    fig,axes = plt.subplots(2,2,figsize=(9,7),layout='constrained')
    mesh = sol.geometry.mesh
    for ax,val,label in zip(axes.ravel(),values,['|Ex|²','|Ey|²','|Ez|²','|E|²']):
        average = np.sum(val*sol.edge_basis.dx,axis=1)/np.sum(sol.edge_basis.dx,axis=1)
        image = ax.tripcolor(mesh.p[0],mesh.p[1],mesh.t.T,facecolors=average,shading='flat',cmap='magma')
        ax.set(aspect='equal',xlabel='x (um)',ylabel='y (um)',title=label)
        if limit:
            ax.set(xlim=(-limit,limit),ylim=(-limit,limit))
            ax.add_patch(plt.Circle((0,0),.3,fill=False,color='cyan',lw=.8))
        fig.colorbar(image,ax=ax,shrink=.8)
    fig.suptitle(title+'  (cell averages, integral |E|² = 1)')
    fig.savefig(OUT/filename,dpi=190);plt.close(fig)


def main():
    start = time.perf_counter()
    data = dict(executed_utc=datetime.now(timezone.utc).isoformat(),python=platform.python_version(),
                versions={p:metadata.version(p) for p in ['numpy','scipy','scikit-fem','gmsh','matplotlib','pytest']})
    rect_levels = []
    refs = rectangle_spectrum()
    for size in (.12,.06,.03):
        sol = rectangle(size)
        rows = []
        for j,(m,(label,exact)) in enumerate(zip(sol.modes,refs)):
            error = None
            if j == 0:
                xy,w,e,_ = sol.quadrature_fields(m)
                f = np.array([np.zeros_like(xy[0]),np.sin(np.pi*xy[0]/1.2),np.zeros_like(xy[0])])
                error = relative_field_error(e,f,w)
            # Do not claim TE/TM identification within the exactly degenerate 11 pair.
            label = label if j < 2 else '11 branch '+('A' if j==2 else 'B')
            rows.append(row(sol,m,label,exact,error))
        rect_levels.append(dict(h_nominal_um=size,**sol.diagnostics,modes=rows))
        print('rectangle',size,'dofs',sol.diagnostics['free_dofs'],'errors',[r['neff_abs_error'] for r in rows],flush=True)
    rect_levels[1]['neff_slopes'] = []
    for j in range(4):
        p = slopes(rect_levels,'h_max_um',[l['modes'][j]['neff_abs_error'] for l in rect_levels])
        for i,pp in enumerate(p):
            rect_levels[i+1]['modes'][j]['neff_slope_from_previous'] = float(pp)
    field_p = slopes(rect_levels,'h_max_um',[l['modes'][0]['field_relative_l2'] for l in rect_levels])
    for i,pp in enumerate(field_p):
        rect_levels[i+1]['modes'][0]['field_slope_from_previous'] = float(pp)
    data['rectangle'] = dict(parameters=dict(n=1.5,width_um=1.2,height_um=.9,wavelength_um=1.55),
                             analytic_te10_area_um2=.72,levels=rect_levels)
    sol.save_npz(OUT/'rectangle_finest_fields.npz')
    field_plot(sol,'rectangle_fields.png','PEC rectangle  TE10')
    loss = rectangle(n=1.5+.001j)
    loss_rows = []
    for m,(label,exact) in zip(loss.modes,rectangle_spectrum(n=1.5+.001j)):
        r = row(loss,m,label,exact)
        r['exact_loss_db_per_m'] = float(20/np.log(10)*(2*np.pi/1.55)*exact.imag*1e6)
        r['relative_loss_error'] = float(abs(m.loss_db_per_m/r['exact_loss_db_per_m']-1))
        loss_rows.append(r)
    data['rectangle_loss'] = dict(material=[1.5,.001],modes=loss_rows)
    ref = FiberReference()
    exact,area = ref.neff,ref.area()
    fiber_levels = []
    for size in (.06,.03,.015):
        sol = fiber(size)
        rows = [row(sol,m,f'HE11 polarization {j+1}',exact,
                    fiber_field_error(sol,m,ref,exact)) for j,m in enumerate(sol.modes)]
        for r in rows:
            r['area_relative_error'] = abs(r['area_um2']/area-1)
        fiber_levels.append(dict(h_nominal_um=size,**sol.diagnostics,modes=rows,
                                 worst_neff_error=max(r['neff_abs_error'] for r in rows)))
        print('fiber',size,'dofs',sol.diagnostics['free_dofs'],'errors',[r['neff_abs_error'] for r in rows],flush=True)
    for key in ['worst_neff_error']:
        p = slopes(fiber_levels,'h_core_max_um',[l[key] for l in fiber_levels])
        for i,pp in enumerate(p):
            fiber_levels[i+1][key+'_slope_actual_core_h'] = float(pp)
    pnom = np.log(np.array([l['worst_neff_error'] for l in fiber_levels[:-1]])/
                  [l['worst_neff_error'] for l in fiber_levels[1:]])/np.log(2)
    for i,pp in enumerate(pnom):
        fiber_levels[i+1]['neff_slope_nominal_h'] = float(pp)
    data['fiber'] = dict(parameters=dict(n_core=1.5,n_clad=1.,core_radius_um=.3,
                            wavelength_um=1.,outer_radius_um=3.,outer_factor=5.),
                        exact_neff=exact,exact_area_um2=area,levels=fiber_levels)
    sol.save_npz(OUT/'fiber_finest_fields.npz')
    field_plot(sol,'fiber_fields.png','Step-index fiber  HE11',limit=1.)
    nc,ns = 1.5+.001j,1.+.0002j
    mesh = sol.geometry
    lossy = solve_vector(mesh,np.where(mesh.region==1,nc,ns),wavelength_um=1.,target_neff=1.25,
                         neff_window=(1.0001,1.4999),candidates=8)
    complex_exact = absorbing_fiber_reference(nc,ns)
    exact_loss = float(20/np.log(10)*2*np.pi*complex_exact.imag*1e6)
    rows = []
    for j,m in enumerate(lossy.modes):
        r = row(lossy,m,f'HE11 polarization {j+1}',complex_exact)
        r.update(exact_loss_db_per_m=exact_loss,relative_loss_error=abs(m.loss_db_per_m/exact_loss-1))
        rows.append(r)
    data['fiber_loss'] = dict(n_core=[nc.real,nc.imag],n_clad=[ns.real,ns.imag],modes=rows)
    lossy.save_npz(OUT/'fiber_absorbing_fields.npz')
    domain = []
    for radius in (2.,2.5,3.):
        s = fiber(h=.02,radius=radius)
        finite = finite_pec_fiber_reference(radius)
        domain.append(dict(outer_radius_um=radius,**s.diagnostics,
            mean_neff=float(np.mean([m.neff.real for m in s.modes])),finite_pec_exact_neff=finite,
            analytic_boundary_error=abs(finite-exact),
            modes=[row(s,m,f'HE11 polarization {j+1}',finite) for j,m in enumerate(s.modes)]))
    data['domain'] = domain
    data['analytic_boundary_only'] = [dict(radius_um=r,finite_neff=finite_pec_fiber_reference(r),
                    error_to_infinite=abs(finite_pec_fiber_reference(r)-exact)) for r in (1.,1.5,2.,2.5,3.)]
    data['elapsed_seconds'] = time.perf_counter()-start
    def serial(o):
        if isinstance(o,np.generic):
            return o.item()
        raise TypeError(type(o))
    (OUT/'validation_m2.json').write_text(json.dumps(data,indent=2,default=serial))
    records = []
    for case,levels in [('rectangle',rect_levels),('fiber',fiber_levels)]:
        for l in levels:
            for r in l['modes']:
                records.append(dict(case=case,h_nominal_um=l['h_nominal_um'],h_max_um=l['h_max_um'],
                              h_core_max_um=l['h_core_max_um'],triangles=l['triangles'],
                              free_dofs=l['free_dofs'],**r))
    keys = list(dict.fromkeys(k for r in records for k in r))
    with (OUT/'convergence_m2.csv').open('w',newline='') as file:
        writer = csv.DictWriter(file,fieldnames=keys);writer.writeheader();writer.writerows(records)
    fig,axes = plt.subplots(1,2,figsize=(10,4),layout='constrained')
    for ax,levels,hkey,title in [(axes[0],rect_levels,'h_max_um','PEC rectangle'),
                                  (axes[1],fiber_levels,'h_core_max_um','Step-index fiber')]:
        h = np.array([l[hkey] for l in levels])
        for j in range(len(levels[0]['modes'])):
            err = np.array([l['modes'][j]['neff_abs_error'] for l in levels])
            ax.loglog(h,err,'o-',label=levels[0]['modes'][j]['label'])
        scale = max(r['neff_abs_error'] for r in levels[0]['modes'])/h[0]**2
        ax.loglog(h,scale*h*h,'k--',alpha=.5,label='O(h²) guide')
        ax.set(xlabel='Actual maximum triangle edge (um)',ylabel='Absolute n_eff error',title=title)
        ax.grid(True,which='both',alpha=.25);ax.legend(fontsize=8)
    fig.savefig(OUT/'convergence_m2.png',dpi=190);plt.close(fig)
    print('reference fiber neff',exact,'area',area,flush=True)
    print('validation elapsed',data['elapsed_seconds'],flush=True)


if __name__ == '__main__':
    main()
