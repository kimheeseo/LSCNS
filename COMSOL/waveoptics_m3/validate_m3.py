"""Run independent analytic benchmarks and report measured mesh convergence."""
from pathlib import Path
from datetime import datetime,timezone
import json,csv,sys,time,platform
import numpy as np
import scipy,skfem,gmsh,matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'tests'))
from analytic_scattering import guide_beta,layer_s,cylinder_field,cylinder_width
from analytic_slab import exact_modes
from waveoptics.scattering import guide_mesh,cylinder_mesh,solve_ports,solve_plane_wave,solve_s_matrix,PML


def pair(z):return [float(np.real(z)),float(np.imag(z))]
def slopes(rows,key):
    return [float(np.log(rows[i][key]/rows[i+1][key])/np.log(rows[i]['h_max_um']/rows[i+1]['h_max_um'])) for i in range(len(rows)-1)]


def main():
    out=ROOT/'results';out.mkdir(exist_ok=True);start=time.perf_counter()
    data=dict(executed_utc=datetime.now(timezone.utc).isoformat(),versions=dict(python=platform.python_version(),numpy=np.__version__,scipy=scipy.__version__,skfem=skfem.__version__,gmsh=gmsh.__version__,matplotlib=matplotlib.__version__),straight={},layers={},cylinder={},pml={},slab={},conversion={})
    th=np.arange(720)*2*np.pi/720;xy=.55*np.array([np.cos(th),np.sin(th)])
    for pol in ['TE','TM']:
        m=1 if pol=='TE' else 0;beta=guide_beta(1.3,1.,.7,m).real;rows=[]
        for h in [.08,.04,.02]:
            s=solve_ports(guide_mesh([0,.9],[0,.7],h),1.3,1.,pol)
            b=s.port_modes['left'].beta[0].real
            coords,w,E,H=s.quadrature_fields();u=E[2] if pol=='TE' else H[2]
            transverse=np.sqrt(2/.7)*np.sin(np.pi*coords[1]/.7) if pol=='TE' else 1.3/np.sqrt(.7)+np.zeros(len(w))
            ref=transverse*np.exp(1j*beta*coords[0]);field_error=float(np.sqrt(np.sum(w*abs(u-ref)**2)/np.sum(w*abs(ref)**2)))
            rows.append(dict(**s.diagnostics,beta_port_per_um=float(b),beta_exact_per_um=float(beta),S11=pair(s.s_parameters['left'][0]),S21=pair(s.s_parameters['right'][0]),transmission_error=float(abs(s.s_parameters['right'][0]-np.exp(1j*beta*.9))),field_l2_error=field_error,reflection_amplitude=float(abs(s.s_parameters['left'][0]))))
        s.save_npz(out/f'straight_{pol}.npz')
        data['straight'][pol]=dict(rows=rows,transmission_slopes=slopes(rows,'transmission_error'),field_slopes=slopes(rows,'field_l2_error'),exact_S21=pair(np.exp(1j*beta*.9)))
        cases={}
        for name,ns,Ls,xb in [('layer',[1.2,1.6,1.2],[.3,.25,.35],[0,.3,.55,.9]),('asymmetric',[1.2,1.6],[.4,.5],[0,.4,.9]),('absorbing',[1.2,1.6+.03j,1.2],[.3,.25,.35],[0,.3,.55,.9])]:
            def index(x,ns=ns,xb=xb):
                idx=np.clip(np.searchsorted(xb,x[0],side='right')-1,0,len(ns)-1);return np.array(ns)[idx]
            s=solve_ports(guide_mesh(xb,[0,.6],.02),index,1.,pol)
            r,t=layer_s(ns,Ls,1.,.6,m,pol)
            cases[name]=dict(FEM_S11=pair(s.s_parameters['left'][0]),exact_S11=pair(r),FEM_S21=pair(s.s_parameters['right'][0]),exact_S21=pair(t),error_S11=float(abs(s.s_parameters['left'][0]-r)),error_S21=float(abs(s.s_parameters['right'][0]-t)),power_out=s.diagnostics['power_out'],exact_power_out=float(abs(r)**2+abs(t)**2),diagnostics=s.diagnostics)
        data['layers'][pol]=cases
        pm=PML((-.8,.8),(-.8,.8),left=.6,right=.6,top=.6,bottom=.6);rows=[]
        ref=cylinder_field(xy,.2,1.5,1.,1.,pol);exactwidth=cylinder_width(.2,1.5,1.,1.,pol)
        for h in [.08,.04,.02]:
            s=solve_plane_wave(cylinder_mesh(.2,pm,h),1.5,1.,1.,pol,pm)
            width=s.scattering_width(.55)
            rows.append(dict(**s.diagnostics,field_l2_error=float(np.linalg.norm(s.sample(xy)-ref)/np.linalg.norm(ref)),scattering_width_um=width,relative_width_error=float(abs(width/exactwidth-1))))
        s.save_npz(out/f'cylinder_{pol}.npz')
        data['cylinder'][pol]=dict(rows=rows,field_slopes=slopes(rows,'field_l2_error'),width_slopes=slopes(rows,'relative_width_error'),exact_scattering_width_um=exactwidth,extraction_radius_um=.55,field_circle_samples=720,width_circle_samples=1440,width_orders=8)
        # Plane-wave fields shown only in physical domain.
        px=np.linspace(-.78,.78,241);XX,YY=np.meshgrid(px,px);points=np.array([XX.ravel(),YY.ravel()]);total=s.sample(points,total=True).reshape(XX.shape)
        np.savez_compressed(out/f'cylinder_plot_{pol}.npz',x=px,total=total)
        controls=[]
        for strength in [0.,1.2,4.]:
            p=PML((0,1.),(0,.7),right=.5,strength=strength)
            v=solve_ports(guide_mesh([0,1.,1.5],[0,.7],.015),1.3,1.,pol,pml=p,ports=('left',))
            exact=-np.exp(2j*beta*1.5-2*beta*strength*.5/3)
            controls.append(dict(strength=strength,FEM_S11=pair(v.s_parameters['left'][0]),exact_S11=pair(exact),amplitude=float(abs(v.s_parameters['left'][0])),exact_amplitude=float(abs(exact)),complex_error=float(abs(v.s_parameters['left'][0]-exact)),**v.diagnostics))
        data['pml'][pol]=controls
        def slabindex(x):return np.where(abs(x[1])<.25,1.5,1.)
        ss=solve_ports(guide_mesh([0,.6],[-2.,-.25,.25,2.],.025),slabindex,1.,pol)
        bs=exact_modes(1.5,1.,.5,1.,pol)[0].beta_per_um
        data['slab'][pol]=dict(exact_beta=float(bs),FEM_port_beta=float(ss.port_modes['left'].beta[0].real),reflection=float(abs(ss.s_parameters['left'][0])),transmission_error=float(abs(ss.s_parameters['right'][0]-np.exp(1j*bs*.6))),S21=pair(ss.s_parameters['right'][0]),**ss.diagnostics)
        def patch(x):return np.where((x[0]>.3)&(x[0]<.55)&(x[1]>.45),1.6,1.2)
        S,channels,diagnostics=solve_s_matrix(guide_mesh([0,.3,.55,.9],[0,.45,.9],.05),patch,1.,pol)
        data['conversion'][pol]=dict(channels=channels,S=[[pair(z) for z in row] for row in S],reciprocity_residual=float(np.linalg.norm(S-S.T)),unitarity_residual=float(np.linalg.norm(S.conj().T@S-np.eye(len(channels)))),conversion_amplitude=float(abs(S[1,0])),max_linear_residual=max(d['linear_residual'] for d in diagnostics))
    # Independent frequency sweep through the dielectric layer, fundamental mode.
    sweep=[]
    mesh=guide_mesh([0,.3,.55,.9],[0,.6],.03)
    def layerindex(x):return np.where((x[0]>.3)&(x[0]<.55),1.6,1.2)
    for pol in ['TE','TM']:
        for wl in np.linspace(.9,1.1,9):
            v=solve_ports(mesh,layerindex,float(wl),pol);r,t=layer_s([1.2,1.6,1.2],[.3,.25,.35],wl,.6,1 if pol=='TE' else 0,pol)
            sweep.append(dict(polarization=pol,wavelength_um=float(wl),reflection_power=float(abs(v.s_parameters['left'][0])**2),transmission_power=float(abs(v.s_parameters['right'][0])**2),exact_reflection_power=float(abs(r)**2),exact_transmission_power=float(abs(t)**2),complex_error_max=float(max(abs(v.s_parameters['left'][0]-r),abs(v.s_parameters['right'][0]-t)))))
    data['sweep']=sweep;data['elapsed_seconds']=time.perf_counter()-start
    (out/'validation_m3.json').write_text(json.dumps(data,indent=2),encoding='utf-8')
    with (out/'convergence_m3.csv').open('w',newline='') as f:
        writer=csv.DictWriter(f,fieldnames=['case','polarization','h_nominal_um','h_max_um','triangles','free_dofs','error','slope_to_next'])
        writer.writeheader()
        for case in ['straight','cylinder']:
            for pol in ['TE','TM']:
                d=data[case][pol];key='transmission_error' if case=='straight' else 'field_l2_error';ps=d['transmission_slopes'] if case=='straight' else d['field_slopes']
                for i,row in enumerate(d['rows']):writer.writerow(dict(case=case,polarization=pol,**{k:row[k] for k in ['h_nominal_um','h_max_um','triangles','free_dofs']},error=row[key],slope_to_next=ps[i] if i<len(ps) else ''))
    fig,axes=plt.subplots(1,2,figsize=(10,4),constrained_layout=True)
    for ax,case,key,label in zip(axes,['straight','cylinder'],['transmission_error','field_l2_error'],['Straight guide complex S21 error','Cylinder scattered-field relative L2 error']):
        for pol in ['TE','TM']:
            rows=data[case][pol]['rows'];ax.loglog([r['h_max_um'] for r in rows],[r[key] for r in rows],'o-',label=pol)
        ax.set(xlabel='Actual maximum edge length [um]',ylabel=label);ax.grid(True,which='both',alpha=.3);ax.legend()
    fig.savefig(out/'convergence_m3.png',dpi=200);plt.close(fig)
    fig,axes=plt.subplots(2,2,figsize=(10,8),constrained_layout=True)
    for i,pol in enumerate(['TE','TM']):
        f=np.load(out/f'cylinder_plot_{pol}.npz');x=f['x'];u=f['total']
        for j,(v,title) in enumerate([(u.real,'Real total scalar field'),(abs(u)**2,'Total scalar intensity')]):
            ax=axes[i,j];im=ax.imshow(v,extent=[x[0],x[-1],x[0],x[-1]],origin='lower',cmap='RdBu_r' if j==0 else 'viridis');fig.colorbar(im,ax=ax);ax.add_patch(plt.Circle((0,0),.2,fill=False,color='black',lw=.8));ax.set(xlabel='x [um]',ylabel='y [um]',title=f'{pol} {title}')
    fig.savefig(out/'cylinder_fields_m3.png',dpi=180);plt.close(fig)
    fig,axes=plt.subplots(1,2,figsize=(10,3.7),constrained_layout=True)
    for ax,pol in zip(axes,['TE','TM']):
        rows=[r for r in sweep if r['polarization']==pol];wl=[r['wavelength_um'] for r in rows]
        for prefix,label,color in [('reflection','R','C0'),('transmission','T','C1')]:
            ax.plot(wl,[r[prefix+'_power'] for r in rows],'o',label=f'FEM {label}',color=color)
            ax.plot(wl,[r['exact_'+prefix+'_power'] for r in rows],'-',label=f'Analytic {label}',color=color)
        ax.set(title=pol+' dielectric layer',xlabel='Wavelength [um]',ylabel='Power fraction');ax.grid(alpha=.3);ax.legend(fontsize=8)
    fig.savefig(out/'spectrum_m3.png',dpi=200);plt.close(fig)
    print(json.dumps({k:data[k] for k in ['executed_utc','versions','elapsed_seconds']},indent=2))


if __name__=='__main__':main()
