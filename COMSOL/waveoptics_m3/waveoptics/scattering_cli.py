"""Reproducible M3 command-line examples; arbitrary media use Python API."""
import argparse,json
from pathlib import Path
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from .scattering import guide_mesh,cylinder_mesh,solve_ports,solve_plane_wave,solve_s_matrix,PML


def pair(z):return [float(np.real(z)),float(np.imag(z))]


def main():
    parser=argparse.ArgumentParser(description='M3 2D TE(Ez)/TM(Hz) FEM scattering; lengths um')
    sub=parser.add_subparsers(dest='case',required=True)
    for name in ['guide','cylinder']:
        p=sub.add_parser(name);p.add_argument('--polarization',choices=['TE','TM'],default='TE')
        p.add_argument('--wavelength',type=float,default=1.);p.add_argument('--h',type=float,default=.04)
        p.add_argument('--out',type=Path,required=True)
        if name=='guide':
            p.add_argument('--width',type=float,default=.6);p.add_argument('--length',type=float,default=.9)
            p.add_argument('--index',type=float,default=1.2);p.add_argument('--layer-index',type=complex)
            p.add_argument('--layer-start',type=float,default=.3);p.add_argument('--layer-end',type=float,default=.55)
            p.add_argument('--incident-mode',type=int,default=0);p.add_argument('--full-s',action='store_true')
        else:
            p.add_argument('--radius',type=float,default=.2);p.add_argument('--index',type=complex,default=1.5+0j)
            p.add_argument('--background-index',type=float,default=1.);p.add_argument('--half-box',type=float,default=.8)
            p.add_argument('--pml-thickness',type=float,default=.6);p.add_argument('--pml-strength',type=float,default=4.)
            p.add_argument('--angle-deg',type=float,default=0.)
            p.add_argument('--extraction-radius',type=float,default=.55)
    args=parser.parse_args();args.out.mkdir(parents=True,exist_ok=True)
    config=vars(args).copy();config['out']=str(config['out'])
    for name,value in list(config.items()):
        if isinstance(value,complex):config[name]=pair(value)
    if args.case=='guide':
        if args.layer_index is None:breaks=[0,args.length];n=args.index
        else:
            if not 0<args.layer_start<args.layer_end<args.length:parser.error('Layer must be strictly inside guide.')
            breaks=[0,args.layer_start,args.layer_end,args.length]
            n=lambda x:np.where((x[0]>args.layer_start)&(x[0]<args.layer_end),args.layer_index,args.index)
        mesh=guide_mesh(breaks,[0,args.width],args.h)
        sol=solve_ports(mesh,n,args.wavelength,args.polarization,incident_mode=args.incident_mode)
        summary=dict(configuration=config,diagnostics=sol.diagnostics,S={side:[pair(z) for z in s] for side,s in sol.s_parameters.items()},port_beta_per_um={side:[pair(z) for z in p.beta[p.propagating]] for side,p in sol.port_modes.items()})
        if args.full_s:
            S,channels,_=solve_s_matrix(mesh,n,args.wavelength,args.polarization)
            np.savez_compressed(args.out/'s_matrix.npz',S=S,channel_port=np.array([c[0] for c in channels]),channel_mode=np.array([c[1] for c in channels]))
            summary.update(S_matrix=[[pair(z) for z in row] for row in S],channels=channels,reciprocity_residual=float(np.linalg.norm(S-S.T)),unitarity_residual=float(np.linalg.norm(S.conj().T@S-np.eye(len(channels)))))
        values=sol.field
    else:
        p=PML((-args.half_box,args.half_box),(-args.half_box,args.half_box),left=args.pml_thickness,right=args.pml_thickness,bottom=args.pml_thickness,top=args.pml_thickness,strength=args.pml_strength)
        mesh=cylinder_mesh(args.radius,p,args.h)
        sol=solve_plane_wave(mesh,args.index,args.background_index,args.wavelength,args.polarization,p,angle_rad=np.deg2rad(args.angle_deg))
        summary=dict(configuration=config,diagnostics=sol.diagnostics,scattering_width_um=sol.scattering_width(args.extraction_radius))
        x=mesh.mesh.p;k=2*np.pi*args.background_index/args.wavelength;angle=np.deg2rad(args.angle_deg)
        values=sol.field+np.exp(1j*k*(np.cos(angle)*x[0]+np.sin(angle)*x[1]))
    sol.save_npz(args.out/'fields.npz')
    (args.out/'summary.json').write_text(json.dumps(summary,indent=2),encoding='utf-8')
    fig,ax=plt.subplots(1,2,figsize=(10,4),constrained_layout=True)
    for a,v,title in zip(ax,[values.real,abs(values)**2],['Real total scalar field','Total scalar intensity']):
        image=a.tripcolor(mesh.mesh.p[0],mesh.mesh.p[1],mesh.mesh.t.T,v,shading='gouraud',cmap='RdBu_r' if 'Real' in title else 'viridis')
        fig.colorbar(image,ax=a);a.set_aspect('equal');a.set(xlabel='x [um]',ylabel='y [um]',title=title)
    fig.savefig(args.out/'field.png',dpi=180);plt.close(fig)
    print(json.dumps(summary,ensure_ascii=False,indent=2))


if __name__=='__main__':main()
