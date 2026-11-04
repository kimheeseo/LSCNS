"""2D SCALAR complex P1 triangular FEM + stretched-coordinate Cartesian PML.

An exploratory FEM research solver for z-invariant silica nested antiresonant
fiber cross-sections, NOT a full-vector Maxwell FEM and NOT validated HCF loss.

Harmonic convention: exp(+i beta z - i omega t); decaying guided/leaky
mode has positive Im(beta). Mesh is approximately interface fitted using
Delaunay boundary samples but is NOT a formally conforming constrained mesh.
Requires mesh, PML and domain convergence before quantitative dB/km claims.
Units: transverse dimensions micrometers, wavelength nanometers.
"""
from __future__ import annotations
from dataclasses import dataclass, asdict, replace
from math import pi, sin, cos, ceil, log
import numpy as np
from scipy.spatial import Delaunay
from scipy.sparse import coo_matrix
from scipy.sparse.linalg import eigs, splu, LinearOperator

C_LIGHT=299792458.0
U01=2.404825557695773
B=(.6961663,.4079426,.8974794)
C=(.0684043**2,.1162414**2,9.896161**2)

def silica_index(wavelength_nm:float)->float:
    w=wavelength_nm/1000.
    if not 220 < wavelength_nm < 3650:
        raise ValueError('Sellmeier model valid for 220 < wavelength_nm < 3650')
    return float(np.sqrt(1+sum(b*w*w/(w*w-c) for b,c in zip(B,C))))

@dataclass(frozen=True)
class HCFDesign:
    coreDiameterUm:float=32.0
    tubeCount:int=6
    nesting:int=2
    outerTubeRatio:float=.78
    innerRatio:float=.61
    thirdRatio:float=.55
    gapCoreUm:float=.25
    minNeighborGapUm:float=.35
    nestedClearanceUm:float=.20
    wall1Um:float=.58
    wall2Um:float=.45
    wall3Um:float=.36
    wavelengthNm:float=1550.
    airIndex:float=1.00027
    autoFit:bool=True
    # FEM configuration: moderate default mesh. Refine for trustworthy trends.
    background_h_um:float=1.80
    wall_edge_h_um:float=.65
    domain_buffer_um:float=5.0
    pml_thickness_um:float=8.0
    pml_strength:float=2.0
    pml_exponent:int=3
    n_modes:int=5
    max_nodes:int=28000

    def validate(self):
        for key in ('coreDiameterUm','outerTubeRatio','innerRatio','thirdRatio','gapCoreUm',
                    'minNeighborGapUm','nestedClearanceUm','wall1Um','wall2Um','wall3Um',
                    'wavelengthNm','airIndex','background_h_um','wall_edge_h_um',
                    'domain_buffer_um','pml_thickness_um','pml_strength'):
            if not np.isfinite(getattr(self,key)):
                raise ValueError(f'Invalid {key}')
        if not(16<=self.coreDiameterUm<=120 and 4<=self.tubeCount<=12
            and isinstance(self.tubeCount,int) and 1<=self.nesting<=3
            and isinstance(self.nesting,int)):
            raise ValueError('Invalid air core, tube count or nesting')
        if not(.2<=self.outerTubeRatio<=1.4 and .3<=self.innerRatio<=.8
               and .3<=self.thirdRatio<=.8):
            raise ValueError('Invalid nested tube ratios')
        if not(0<=self.gapCoreUm<=6 and 0<=self.minNeighborGapUm<=6
               and 0<=self.nestedClearanceUm<=1.5):
            raise ValueError('Invalid clearances')
        if not(all(.10<=x<=1.6 for x in (self.wall1Um,self.wall2Um,self.wall3Um))):
            raise ValueError('Invalid silica wall thickness')
        if not(1<=self.airIndex<=1.01 and .25<=self.background_h_um<=8
               and .10<=self.wall_edge_h_um<=2.5
               and 2<=self.domain_buffer_um<=40
               and 2<=self.pml_thickness_um<=40
               and 0.05<=self.pml_strength<=12
               and 1<=self.pml_exponent<=6 and 2<=self.n_modes<=12
               and 1000<=self.max_nodes<=250000):
            raise ValueError('Invalid numerical FEM settings')
        silica_index(self.wavelengthNm)
        return self


def design_from_dict(d:dict)->HCFDesign:
    keys=HCFDesign.__dataclass_fields__.keys()
    safe={key:d[key] for key in keys if key in d}
    if 'tubeCount' in safe: safe['tubeCount']=int(safe['tubeCount'])
    if 'nesting' in safe: safe['nesting']=int(safe['nesting'])
    if 'n_modes' in safe: safe['n_modes']=int(safe['n_modes'])
    if 'pml_exponent' in safe: safe['pml_exponent']=int(safe['pml_exponent'])
    if 'max_nodes' in safe: safe['max_nodes']=int(safe['max_nodes'])
    return HCFDesign(**safe).validate()


def geometry(d:HCFDesign)->dict:
    d.validate()
    coreR=d.coreDiameterUm/2
    s=sin(pi/d.tubeCount)
    requestedD=d.coreDiameterUm*d.outerTubeRatio
    maxD=(2*s*(coreR+d.gapCoreUm)-d.minNeighborGapUm)/(1-s)
    actualD=min(requestedD,maxD) if d.autoFit else requestedD
    if actualD<=0:raise ValueError('Tube overlap / negative allowed diameter')
    r1=actualD/2
    radii=[r1,r1*d.innerRatio,r1*d.innerRatio*d.thirdRatio][:d.nesting]
    walls=[d.wall1Um,d.wall2Um,d.wall3Um][:d.nesting]
    offsets=[0.]
    for i in range(1,d.nesting):
        off=radii[i-1]-walls[i-1]-radii[i]-d.nestedClearanceUm
        if off<0:raise ValueError('Nested tube does not fit in parent tube')
        offsets.append(offsets[-1]+off)
    for r,w in zip(radii,walls):
        if w>=.8*r:raise ValueError('Tube wall thicker than allowed')
    ringR=coreR+r1+d.gapCoreUm
    gap=2*ringR*s-actualD
    if gap<d.minNeighborGapUm-1e-7:raise ValueError('Adjacent outer tubes overlap')
    tubes=[]
    for i in range(d.tubeCount):
        angle=2*pi*i/d.tubeCount-pi/2
        for j,(r,w,off) in enumerate(zip(radii,walls,offsets)):
            tubes.append(dict(tube=i+1,layer=j+1,x=ringR*cos(angle)+off*cos(angle),
                              y=ringR*sin(angle)+off*sin(angle),r=r,wall=w))
    extent=max(np.hypot(t['x'],t['y'])+t['r'] for t in tubes)
    return dict(core_radius_um=coreR,outer_tube_diameter_um=actualD,
                requested_outer_diameter_um=requestedD,neighbor_gap_um=gap,
                radius_extent_um=float(extent),tubes=tubes,
                auto_fitted=bool(actualD<requestedD-1e-7))


def is_silica(x:np.ndarray,y:np.ndarray,tubes:list)->np.ndarray:
    x=np.asarray(x);y=np.asarray(y)
    mask=np.zeros(np.broadcast_shapes(x.shape,y.shape),dtype=bool)
    for t in tubes:
        rr=(x-t['x'])**2+(y-t['y'])**2
        mask|=(rr < t['r']**2)&(rr > (t['r']-t['wall'])**2)
    return mask


def mesh_geometry(d:HCFDesign,g:dict|None=None):
    """Adaptive Delaunay P1 triangles with explicit wall interface samples.
    This is NOT a formally constrained/conforming mesh. Mesh uncertainty must
    be checked, especially near thin walls, nodes and antiresonance peaks.
    """
    g=geometry(d) if g is None else g
    window=g['radius_extent_um']+d.domain_buffer_um
    half=window+d.pml_thickness_um
    h=d.background_h_um
    side=np.linspace(-half,half,max(9,int(ceil(2*half/h))+1))
    XX,YY=np.meshgrid(side,side,indexing='ij')
    px=XX.ravel();py=YY.ravel()
    # Remove background points very close to a glass boundary to avoid needles.
    keep=np.ones(len(px),bool)
    for tube in g['tubes']:
        radius=np.hypot(px-tube['x'],py-tube['y'])
        keep &= (np.abs(radius-tube['r'])>.32*h)&(np.abs(radius-(tube['r']-tube['wall']))>.32*h)
    points=[np.column_stack((px[keep],py[keep]))]
    # Explicit circular interface points on both wall boundaries and within wall.
    for t in g['tubes']:
        # Additional through-wall rings create triangles with nonzero silica volume.
        for radius in (t['r'],t['r']-.5*t['wall'],t['r']-t['wall']):
            npts=max(24,int(ceil(2*pi*radius/d.wall_edge_h_um)))
            ang=np.arange(npts)*2*pi/npts
            points.append(np.column_stack((t['x']+radius*np.cos(ang),
                                           t['y']+radius*np.sin(ang))))
    # Ensure outer Dirichlet square has coordinates on each side.
    borders=np.linspace(-half,half,max(21,int(ceil(2*half/h))+1))
    points.extend([np.column_stack((borders,np.full_like(borders,half))),
                   np.column_stack((borders,np.full_like(borders,-half))),
                   np.column_stack((np.full_like(borders,half),borders)),
                   np.column_stack((np.full_like(borders,-half),borders))])
    xy=np.unique(np.round(np.concatenate(points),8),axis=0)
    if len(xy)>d.max_nodes:raise ValueError(f'Mesh nodes {len(xy)} exceed max_nodes {d.max_nodes}; increase mesh size')
    tris=Delaunay(xy).simplices
    # Keep only square interior and positive-area elements.
    p=xy[tris]
    e1=p[:,1]-p[:,0];e2=p[:,2]-p[:,0]
    cr=e1[:,0]*e2[:,1]-e1[:,1]*e2[:,0]
    valid=np.abs(cr)>1e-10
    tris=tris[valid]
    boundary=(np.isclose(np.abs(xy[:,0]),half,atol=1e-7)|
              np.isclose(np.abs(xy[:,1]),half,atol=1e-7))
    return xy,tris,boundary,window,half


def assemble(d:HCFDesign,return_mesh=False):
    g=geometry(d)
    xy,tri,boundary,window,half=mesh_geometry(d,g)
    p=xy[tri]
    cross=(p[:,1,0]-p[:,0,0])*(p[:,2,1]-p[:,0,1])-(p[:,2,0]-p[:,0,0])*(p[:,1,1]-p[:,0,1])
    area=np.abs(cross)/2
    gx=np.stack([p[:,1,1]-p[:,2,1],p[:,2,1]-p[:,0,1],p[:,0,1]-p[:,1,1]],axis=1)/cross[:,None]
    gy=np.stack([p[:,2,0]-p[:,1,0],p[:,0,0]-p[:,2,0],p[:,1,0]-p[:,0,0]],axis=1)/cross[:,None]
    xm=p[:,:,0].mean(axis=1);ym=p[:,:,1].mean(axis=1)
    rx=np.maximum((np.abs(xm)-window)/d.pml_thickness_um,0)
    ry=np.maximum((np.abs(ym)-window)/d.pml_thickness_um,0)
    sx=1+1j*d.pml_strength*rx**d.pml_exponent
    sy=1+1j*d.pml_strength*ry**d.pml_exponent
    b=sx*sy
    n=np.where(is_silica(xm,ym,g['tubes']),silica_index(d.wavelengthNm),d.airIndex)
    k0=2*pi/(d.wavelengthNm/1000)
    mass=(area*b)[:,None,None]*(np.ones((3,3))+np.eye(3))[None,:,:]/12
    kmat=area[:,None,None]*((sy/sx)[:,None,None]*gx[:,:,None]*gx[:,None,:]
                              +(sx/sy)[:,None,None]*gy[:,:,None]*gy[:,None,:])
    a=(k0*k0*n*n)[:,None,None]*mass-kmat
    rows=np.broadcast_to(tri[:,:,None],(len(tri),3,3)).ravel()
    cols=np.broadcast_to(tri[:,None,:],(len(tri),3,3)).ravel()
    N=len(xy)
    A=coo_matrix((a.ravel(),(rows,cols)),shape=(N,N)).tocsr()
    M=coo_matrix((mass.ravel(),(rows,cols)),shape=(N,N)).tocsr()
    keep=np.flatnonzero(~boundary)
    # Lumped area weights for physical-region modal integrals (no PML).
    physical=np.maximum(np.abs(xm),np.abs(ym))<window
    lump=np.bincount(tri[physical].ravel(),weights=np.repeat(area[physical]/3,3),minlength=N)
    core=np.hypot(xy[:,0],xy[:,1]) < g['core_radius_um']
    if return_mesh:return A[keep,:][:,keep],M[keep,:][:,keep],xy[keep],lump[keep],core[keep],dict(nodes=N,elements=len(tri),window_um=window,halfwidth_um=half),xy,tri
    return A[keep,:][:,keep],M[keep,:][:,keep],xy[keep],lump[keep],core[keep],dict(nodes=N,elements=len(tri),window_um=window,halfwidth_um=half)


def ideal_neff(d:HCFDesign)->float:
    lam=d.wavelengthNm/1000
    return float(np.sqrt(d.airIndex*d.airIndex-(U01*lam/(pi*d.coreDiameterUm))**2))


def solve(d:HCFDesign,include_field=False):
    import warnings
    d.validate()
    A,M,xy,lump,core,diag,*mesh=assemble(d,include_field)
    k0=2*pi/(d.wavelengthNm/1000)
    target=ideal_neff(d)
    sigma=(k0*target)**2
    lu=splu((A-sigma*M).tocsc())
    operator=LinearOperator(A.shape,dtype=complex,matvec=lambda v:lu.solve(M@v))
    mu,vectors=eigs(operator,k=d.n_modes,which='LM',tol=5e-8,maxiter=1300)
    neff=np.sqrt(sigma+1/mu)/k0
    # Select the mode with largest fundamental J0 overlap inside the air core.
    rr=np.hypot(xy[:,0],xy[:,1]);from scipy.special import j0
    reference=np.where(core,j0(U01*rr/ (d.coreDiameterUm/2)),0.)
    scores=[]
    for i in range(len(neff)):
        vec=vectors[:,i]
        w=np.maximum(lump,0)
        pn=float(np.sum(w*np.abs(vec)**2))
        cn=float(np.sum(w[core]*np.abs(vec[core])**2))
        refn=float(np.sum(w*reference*reference))
        corr=float(np.abs(np.sum(w*np.conj(reference)*vec))**2/max(refn*pn,1e-40))
        scores.append((corr,cn/max(pn,1e-40),i))
    scores.sort(reverse=True)
    idx=scores[0][2]
    ne=neff[idx]
    # exp(+i beta z) -> Im(beta)>0 means attenuation. Nonpositive imag means unreliable sign/PML result.
    attenuation_db_km=40*pi*1000*float(ne.imag)/(d.wavelengthNm/1e9*log(10)) if ne.imag>0 else None
    # Note wavelength_nm/1e9 = metres.
    mode=vectors[:,idx]
    intensity=np.abs(mode)**2
    core_overlap=float(np.sum(lump[core]*intensity[core])/max(np.sum(lump*intensity),1e-30))
    out=dict(model='2D_scalar_complex_P1_triangular_FEM_Cartesian_PML',
             full_vector=False,converged_and_validated=False,
             input=asdict(d),geometry=geometry(d),
             neff_real=float(ne.real),neff_imag=float(ne.imag),
             neff_ideal_capillary=target,confinement_loss_dB_km_unverified=attenuation_db_km,
             air_core_power_fraction=core_overlap,
             ideal_J0_overlap=scores[0][0],mode_identification_pass=bool(scores[0][0]>=.75),mesh=diag,
             ranked_modes=[dict(neff_real=float(neff[i].real),neff_imag=float(neff[i].imag),
                                J0_overlap=float(corr),core_fraction=float(frac))
                           for corr,frac,i in scores],
             limitations=['SCALAR FEM; HE11 full-vector boundary continuity is not resolved',
                          'Delaunay point mesh is not formally interface-conforming',
                          'PML and mesh convergence not established by this single solve',
                          'Loss dB/km is preliminary and must not be treated as measured/validated',
                          'Resonance features must be extracted from FEM loss spectrum, not only ARROW formula'])
    if include_field:
        # Export subsampled field point cloud, not a fabricated Gaussian profile.
        amp=mode/max(float(np.max(np.abs(mode))),1e-30)
        mask=(np.max(np.abs(xy),axis=1)<diag['window_um'])
        order=np.flatnonzero(mask)
        if len(order)>5000:order=order[::int(ceil(len(order)/5000))]
        out['mode_points']=[dict(x_um=float(xy[j,0]),y_um=float(xy[j,1]),
                                 amplitude_real=float(amp[j].real),amplitude_imag=float(amp[j].imag),
                                 intensity=float(abs(amp[j])**2)) for j in order]
    return out


def peaks_from_fem_sweep(results:list):
    """Only local extrema of calculated FEM attenuation. No invented ARROW peak."""
    r=sorted(results,key=lambda q:q['input']['wavelengthNm'])
    out=[]
    for a,b,c in zip(r[:-2],r[1:-1],r[2:]):
        if not all(x.get('mode_identification_pass') for x in (a,b,c)):continue
        vals=[x.get('confinement_loss_dB_km_unverified') for x in (a,b,c)]
        if any(v is None or v<=0 for v in vals):continue
        if vals[1]>vals[0] and vals[1]>vals[2]:kind='resonance_loss_peak_candidate'
        elif vals[1]<vals[0] and vals[1]<vals[2]:kind='antiresonance_loss_dip_candidate'
        else:continue
        out.append(dict(kind=kind,wavelength_nm=b['input']['wavelengthNm'],
                        frequency_THz=C_LIGHT/(b['input']['wavelengthNm']*1e-9)/1e12,
                        fem_loss_dB_km_unverified=vals[1],verified=False,
                        label='FEM-grid local extremum only; requires fine scan/PML convergence'))
    return out


def sweep(d:HCFDesign,wavelengths_nm,include_field_for_center=False):
    arr=[];lams=sorted(set(float(x) for x in wavelengths_nm))
    field_index=min(range(len(lams)),key=lambda i:abs(lams[i]-d.wavelengthNm)) if lams else -1
    for i,wl in enumerate(lams):
        cfg=replace(d,wavelengthNm=wl)
        arr.append(solve(cfg,include_field=include_field_for_center and i==field_index))
    return dict(solver='scalar_complex_FEM_PML_unverified',results=arr,
                extrema_candidates=peaks_from_fem_sweep(arr))


def convergence_check(d:HCFDesign):
    base=solve(d)
    finer=solve(replace(d,wall_edge_h_um=d.wall_edge_h_um*.78,background_h_um=d.background_h_um*.85))
    thicker=solve(replace(d,pml_thickness_um=d.pml_thickness_um*1.30))
    def difference(a,b):
        x=a['confinement_loss_dB_km_unverified'];y=b['confinement_loss_dB_km_unverified']
        return None if x is None or y is None or y<=0 else abs(x-y)/abs(y)*100
    e_mesh=difference(base,finer);e_pml=difference(base,thicker)
    pass_loss=bool(e_mesh is not None and e_mesh<5 and e_pml is not None and e_pml<5
                   and all(x['mode_identification_pass'] for x in (base,finer,thicker)))
    return dict(base=base,mesh=finer,pml=thicker,
                relative_mesh_change_pct=e_mesh,relative_pml_change_pct=e_pml,
                preliminary_loss_convergence_pass=pass_loss,
                note='Even 5% 2-point convergence is insufficient for publication-grade HCF confinement loss; domain and mode checks also required.')

if __name__=='__main__':
 import argparse,json
 from pathlib import Path
 ap=argparse.ArgumentParser()
 ap.add_argument('--design-json',type=Path)
 ap.add_argument('--output',type=Path,default=Path('hcf_fem_result.json'))
 ap.add_argument('--sweep',nargs='*',type=float)
 ap.add_argument('--field',action='store_true')
 args=ap.parse_args()
 conf=design_from_dict(json.loads(args.design_json.read_text()) if args.design_json else {})
 result=sweep(conf,args.sweep,include_field_for_center=args.field) if args.sweep else solve(conf,include_field=args.field)
 args.output.parent.mkdir(parents=True,exist_ok=True)
 args.output.write_text(json.dumps(result,indent=2),encoding='utf-8')
 print(f'Wrote {args.output}')
