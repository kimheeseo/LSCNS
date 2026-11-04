"""Solid silica fiber scalar paraxial (2+1)D split-step FFT beam propagation.
Real 3D intensity slices E(x,y,z) calculated by longitudinal marching.
NOT 3D Full-vector Maxwell/FEM/FDTD. HCF/NANF leaky modes are NOT supported.
x/y/z in micrometres; wavelength in nanometres.
"""
from __future__ import annotations
from dataclasses import dataclass, asdict
import numpy as np

_B=(.6961663,.4079426,.8974794)
_C=(.0684043**2,.1162414**2,9.896161**2)

def silica_index(wavelength_nm):
    l=wavelength_nm/1000.
    if not .21<l<3.7: raise ValueError("Silica Sellmeier outside 210–3700 nm")
    return float(np.sqrt(1+sum(b*l*l/(l*l-c) for b,c in zip(_B,_C))))

@dataclass
class Design:
    profile:str='step' # step, graded, trench, multilayer, homogeneous
    wavelength_nm:float=1550.
    core_radius_um:float=5.5
    delta_n_core:float=.0034785793846155
    graded_alpha:float=2.
    trench_start_um:float=7.
    trench_end_um:float=10.
    trench_delta_n:float=-.001
    shells:tuple=((4.,.006),(5.,.002),(7.,0.),(9.,-.0008))
    half_width_um:float=36.
    grid:int=128
    length_um:float=300.
    steps:int=120
    waist_um:float=4.5
    offset_um:float=1.5
    snapshots:int=9

def index_map(d,X,Y):
    radius=np.hypot(X,Y);n=silica_index(d.wavelength_nm)
    if d.profile=='homogeneous':return np.full_like(radius,n)
    if d.profile=='step':dn=np.where(radius<d.core_radius_um,d.delta_n_core,0.)
    elif d.profile=='graded':dn=np.where(radius<d.core_radius_um,d.delta_n_core*np.maximum(1-(radius/d.core_radius_um)**d.graded_alpha,0.),0.)
    elif d.profile=='trench':dn=np.where(radius<d.core_radius_um,d.delta_n_core,np.where((radius>=d.trench_start_um)&(radius<d.trench_end_um),d.trench_delta_n,0.))
    elif d.profile=='multilayer':
        dn=np.zeros_like(radius);prev=0
        for boundary,value in d.shells:dn=np.where((radius>=prev)&(radius<boundary),value,dn);prev=boundary
    else:raise ValueError("Unsupported profile; HCF geometry requires complex full-vector/PML solver")
    return n+dn

def metrics(E,X,Y,dx,halfwidth):
    I=np.abs(E)**2;power=float(I.sum()*dx*dx)
    xc=float(np.sum(X*I)*dx*dx/power);yc=float(np.sum(Y*I)*dx*dx/power)
    variance=float(np.sum(((X-xc)**2+(Y-yc)**2)*I)*dx*dx/power)
    aeff=float(power*power/(np.sum(I*I)*dx*dx))
    border=(np.abs(X)>.85*halfwidth)|(np.abs(Y)>.85*halfwidth)
    edge=float(np.sum(I[border])*dx*dx/power)
    return dict(power=power,centroid_x_um=xc,centroid_y_um=yc,
                MFD_2nd_um=2*np.sqrt(2*variance),Aeff_um2=aeff,edge_power_fraction=edge)

def propagate(d=Design()):
    if d.profile not in ('step','graded','trench','multilayer','homogeneous'):
        raise ValueError('HCF not supported by scalar BPM')
    if not isinstance(d.grid,int) or not 32<=d.grid<=512 or d.grid&(d.grid-1):
        raise ValueError('Grid must be power of two from 32 to 512')
    if not isinstance(d.steps,int) or not 1<=d.steps<=2000 or not 2<=d.snapshots<=d.steps+1:
        raise ValueError('Invalid steps/snapshots')
    if not np.isfinite(d.half_width_um) or d.half_width_um < max(2*d.core_radius_um,4*d.waist_um):
        raise ValueError('Enlarge periodic boundary window')
    if not (d.length_um>0 and d.waist_um>0 and np.isfinite(d.offset_um)):
        raise ValueError('Invalid geometry / launch')
    if d.profile=='trench' and not(d.core_radius_um<d.trench_start_um<d.trench_end_um<d.half_width_um):
        raise ValueError('Invalid trench dimensions')
    if d.profile=='multilayer':
        prev=0
        for r,dn in d.shells:
            if not(prev<r<d.half_width_um and np.isfinite(dn)):raise ValueError('Invalid shells')
            prev=r
    n=d.grid;dx=2*d.half_width_um/n
    a=(np.arange(n)-n//2)*dx;X,Y=np.meshgrid(a,a,indexing='xy')
    nm=index_map(d,X,Y);nref=silica_index(d.wavelength_nm);k0=2*np.pi/(d.wavelength_nm/1000)
    dz=d.length_um/d.steps
    q=2*np.pi*np.fft.fftfreq(n,d=dx);KX,KY=np.meshgrid(q,q,indexing='xy')
    halfphase=np.exp(1j*dz*k0*(nm*nm-nref*nref)/(4*nref))
    diffraction=np.exp(-1j*dz*(KX*KX+KY*KY)/(2*k0*nref))
    E=np.exp(-((X-d.offset_um)**2+Y*Y)/d.waist_um**2).astype(complex)
    E/=np.sqrt(np.sum(np.abs(E)**2)*dx*dx)
    mark=set(np.rint(np.linspace(0,d.steps,d.snapshots)).astype(int))
    history=[];snaps=[]
    for j in range(d.steps+1):
        row=dict(z_um=j*dz,**metrics(E,X,Y,dx,d.half_width_um));history.append(row)
        if j in mark:snaps.append(dict(z_um=j*dz,intensity=np.abs(E)**2))
        if j<d.steps:
            E*=halfphase;E=np.fft.ifft2(np.fft.fft2(E)*diffraction);E*=halfphase
    return dict(design=asdict(d),axis_um=a,index_map=nm,history=history,slices=snaps,
                max_power_error=max(abs(m['power']-1) for m in history),
                HCF_loss_dB_km=None,
                solver='2+1D scalar paraxial FFT BPM; no 3D vector FEM, no PML',
                boundary='periodic FFT; check boundary energy and convergence')

def export(r,folder='results'):
    from pathlib import Path
    import json,csv
    p=Path(folder);p.mkdir(parents=True,exist_ok=True)
    fields=list(r['history'][0]); 
    with (p/'propagation_metrics.csv').open('w',newline='') as f:
        writer=csv.DictWriter(f,fieldnames=fields);writer.writeheader();writer.writerows(r['history'])
    summary={k:v for k,v in r.items() if k not in ('axis_um','index_map','slices','history')}
    summary['final']=r['history'][-1]
    (p/'propagation_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2))
    np.savez_compressed(p/'field_volume_slices.npz',x=r['axis_um'],index=r['index_map'],
          z=np.array([s['z_um'] for s in r['slices']]),intensity=np.array([s['intensity'] for s in r['slices']]))
    return summary

if __name__=='__main__':
    r=propagate();print(export(r))
