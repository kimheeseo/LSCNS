"""2D scalar P1 triangular FEM with complex PML for optical-fiber bend modes.

Research prototype inspired by Velamuri et al. (COMSOL Conference 2018).
This is a real 2-D eigenmode numerical solve, NOT their full-vector EWFD/GEBT,
nor a verified model of any commercial cable or actual physical coating.
The PML is on a shortened glass-only computational window: do NOT interpret
its leakage rate as certified macrobend loss without mesh/PML/window sweeps.

Equation: div(C grad E) + k0^2 n_eq^2 s_x s_y E = beta^2 s_x s_y E
C=diag(s_y/s_x, s_x/s_y), n_eq(x,y)=n_stress(r)(1+x/R).
Outer boundary is zero after the complex-coordinate-stretched PML.
"""
from __future__ import annotations
from dataclasses import dataclass, asdict, replace
from math import pi, log
import numpy as np
from scipy.sparse import coo_matrix
from scipy.sparse.linalg import eigs, splu, LinearOperator

_B = (.6961663,.4079426,.8974794)
_C = (.0684043**2,.1162414**2,9.896161**2)

def silica_index(wavelength_nm):
    w = wavelength_nm / 1000.
    if not (.21 < w < 3.7):
        raise ValueError("silica Sellmeier wavelength requires 210-3700 nm")
    return np.sqrt(1 + sum(b*w*w/(w*w-c) for b,c in zip(_B,_C)))

@dataclass(frozen=True)
class Fiber:
    wavelength_nm:float = 1550.
    core_radius_um:float = 4.1
    delta_n_core:float = .005
    profile:str = "step"
    trench_start_um:float = 7.
    trench_end_um:float = 12.
    trench_delta_n:float = -.002
    dip_radius_um:float = 0.
    dip_delta_n:float = 0.
    bend_diameter_mm:float = 10.
    stress_mode:str = "none"
    poisson_ratio:float = .17
    p11:float = .113
    p12:float = .252
    window_um:float = 22.
    pml_thickness_um:float = 10.85
    pml_strength:float = 2.
    pml_exponent:int = 3
    grid_h_um:float = .7
    n_modes:int = 10
    max_points:int = 220

def validate(f):
    if f.profile not in ("step","trench"):
        raise ValueError("profile must be step or trench")
    if f.stress_mode not in ("none","beam_proxy"):
        raise ValueError("stress_mode must be none or beam_proxy (NOT GEBT)")
    for key in ("wavelength_nm","core_radius_um","delta_n_core","bend_diameter_mm",
                "window_um","pml_thickness_um","pml_strength","grid_h_um"):
        if not np.isfinite(getattr(f,key)):
            raise ValueError("invalid "+key)
    if f.core_radius_um <= 0 or f.delta_n_core <= 0 or f.window_um <= 2*f.core_radius_um:
        raise ValueError("need positive core/index step and window >2*core radius")
    if not 0 <= f.bend_diameter_mm <= 1000:
        raise ValueError("bend diameter must be 0..1000 mm, zero is straight")
    if min(f.pml_thickness_um,f.pml_strength,f.grid_h_um) <= 0:
        raise ValueError("PML thickness/strength and mesh size must be positive")
    if not 2 <= f.n_modes <= 15:
        raise ValueError("n_modes must be 2..15")
    if f.stress_mode == "beam_proxy" and not 0 <= f.poisson_ratio < .5:
        raise ValueError("invalid Poisson ratio")
    if f.profile == "trench" and not (
        f.core_radius_um < f.trench_start_um < f.trench_end_um < f.window_um
    ):
        raise ValueError("trench must lie inside the non-PML computational window")
    if f.bend_diameter_mm and (
        f.window_um + f.pml_thickness_um >= 500*f.bend_diameter_mm
    ):
        raise ValueError("2D computational half-width must be smaller than bend radius")
    silica_index(f.wavelength_nm)
    return f

def index_field(f,x,y):
    r=np.hypot(x,y)
    n=silica_index(f.wavelength_nm) + f.delta_n_core*(r < f.core_radius_um)
    if f.profile=="trench":
        n=n+f.trench_delta_n*((r>=f.trench_start_um)&(r<f.trench_end_um))
    if f.dip_radius_um>0:
        n=n+f.dip_delta_n*(r<f.dip_radius_um)
    if f.bend_diameter_mm>0:
        eps=x/(f.bend_diameter_mm*500.)
        if f.stress_mode=="beam_proxy":
            n=n*(1-.5*n*n*(f.p11*eps-2*f.p12*f.poisson_ratio*eps))
        n=n*(1+eps)
    return n

def stretch(f,x):
    frac=np.maximum((np.abs(x)-f.window_um)/f.pml_thickness_um,0)
    return 1+1j*f.pml_strength*frac**f.pml_exponent

def assemble(f):
    validate(f)
    half=f.window_um+f.pml_thickness_um
    nx=max(11,int(np.ceil(2*half/f.grid_h_um))+1)
    if nx%2==0: nx+=1
    if nx>f.max_points:
        raise ValueError("mesh too large; increase grid_h_um or max_points")
    xx=np.linspace(-half,half,nx)
    h=float(xx[1]-xx[0])
    i,j=np.meshgrid(np.arange(nx-1),np.arange(nx-1),indexing="ij")
    v=(i*nx+j).ravel()
    triangles=np.vstack([
        np.stack([v,v+nx,v+nx+1],axis=1),
        np.stack([v,v+nx+1,v+1],axis=1)])
    nt=len(triangles)
    X,Y=np.meshgrid(xx,xx,indexing="ij")
    coords=np.stack([X.ravel(),Y.ravel()],axis=1)
    pts=coords[triangles]
    x,y=pts[:,:,0],pts[:,:,1]
    cross=((x[:,1]-x[:,0])*(y[:,2]-y[:,0])-
           (x[:,2]-x[:,0])*(y[:,1]-y[:,0]))
    area=np.abs(cross)/2
    gx=np.stack([y[:,1]-y[:,2],y[:,2]-y[:,0],y[:,0]-y[:,1]],axis=1)/cross[:,None]
    gy=np.stack([x[:,2]-x[:,1],x[:,0]-x[:,2],x[:,1]-x[:,0]],axis=1)/cross[:,None]
    xc,yc=x.mean(axis=1),y.mean(axis=1)
    sx,sy=stretch(f,xc),stretch(f,yc)
    b=sx*sy
    k0=2*pi/(f.wavelength_nm/1000)
    n=index_field(f,xc,yc)
    M=(area*b)[:,None,None]*(np.ones((3,3))+np.eye(3))[None,:,:]/12
    A=(k0*k0*n*n)[:,None,None]*M-area[:,None,None]*(
        (sy/sx)[:,None,None]*gx[:,:,None]*gx[:,None,:]+
        (sx/sy)[:,None,None]*gy[:,:,None]*gy[:,None,:])
    nodes=nx*nx
    rows=np.broadcast_to(triangles[:,:,None],(nt,3,3)).ravel()
    cols=np.broadcast_to(triangles[:,None,:],(nt,3,3)).ravel()
    matA=coo_matrix((A.ravel(),(rows,cols)),shape=(nodes,nodes)).tocsr()
    matM=coo_matrix((M.ravel(),(rows,cols)),shape=(nodes,nodes)).tocsr()
    idx=np.arange(nx)[:,None]
    jdx=np.arange(nx)[None,:]
    interior=(idx>0)&(idx<nx-1)&(jdx>0)&(jdx<nx-1)
    keep=np.flatnonzero(interior.ravel())
    return matA[keep,:][:,keep],matM[keep,:][:,keep],coords[keep],nx,h

def solve(f):
    A,M,coords,nx,h=assemble(f)
    k0=2*pi/(f.wavelength_nm/1000)
    n_guess=silica_index(f.wavelength_nm)+.45*f.delta_n_core
    sigma=(k0*n_guess)**2
    # For complex PML mass, SciPy eigs(A,M=M,sigma) is not safe: M is
    # non-Hermitian/non-SPD. Apply generalized shift-invert explicitly.
    lu=splu((A-sigma*M).tocsc())
    operator=LinearOperator(A.shape,dtype=complex,matvec=lambda v:lu.solve(M@v))
    mu,vecs=eigs(operator,k=f.n_modes,which="LM",tol=1e-9,maxiter=1500)
    beta_sq=sigma+1/mu
    neffs=np.sqrt(beta_sq)/k0
    r=np.hypot(coords[:,0],coords[:,1])
    actual=(np.abs(coords).max(axis=1)<f.window_um)
    nearcore=(r<2*f.core_radius_um)
    modes=[]
    for idx,neff in enumerate(neffs):
        power=np.abs(vecs[:,idx])**2
        core_fraction=float(power[nearcore].sum()/max(power[actual].sum(),1e-30))
        physical_fraction=float(power[actual].sum()/max(power.sum(),1e-30))
        modes.append(dict(index=idx,neff_real=float(neff.real),
            neff_imag=float(neff.imag),core_fraction=core_fraction,
            physical_fraction=physical_fraction))
    selected=max(modes,key=lambda mode:mode["core_fraction"])
    imag=selected["neff_imag"]
    alpha_db_m=20/log(10)*k0*1e6*imag
    bend_db_turn=(alpha_db_m*pi*f.bend_diameter_mm/1000
                  if f.bend_diameter_mm>0 and imag>=0 else None)
    return dict(
        input=asdict(f),neff_real=selected["neff_real"],neff_imag=imag,
        bend_loss_dB_per_turn=bend_db_turn,loss_dB_per_m=alpha_db_m,
        core_fraction=selected["core_fraction"],
        physical_fraction=selected["physical_fraction"],
        grid_n=nx,h_um=h,unknowns=A.shape[0],ranked_modes=modes,
        nonnegative_imag=bool(imag>=0),is_fully_validated=False,
        solver_level="2D scalar P1 triangular FEM with complex PML (NOT full-vector COMSOL/GEBT)",
        limitations=[
            "Square PML truncates actual 125um cladding; coating not modeled",
            "Requires independent PML thickness, mesh, domain and mode-tracking convergence",
            "Optional linear beam/Poisson stress proxy is NOT GEBT",
            "No experimental product or paper profile-level validation"])

def convergence_study(f,mesh_values=(.8,.65,.55),
                      pml_values=(6.2,9.3,10.85),tolerance_percent=2.):
    """Real independent FEM reruns; not an assertion of convergence."""
    base=solve(f)
    mesh=[dict(grid_h_requested_um=h,**solve(replace(f,grid_h_um=h)))
          for h in mesh_values]
    pml=[dict(pml_thickness_requested_um=t,
              **solve(replace(f,pml_thickness_um=t))) for t in pml_values]
    def compare(a,b):
        x=a["bend_loss_dB_per_turn"];y=b["bend_loss_dB_per_turn"]
        if x is None or y is None or y<=0:return None
        return 100*abs(x-y)/abs(y)
    mesh_change=compare(mesh[-2],mesh[-1]) if len(mesh)>=2 else None
    pml_change=compare(pml[-2],pml[-1]) if len(pml)>=2 else None
    passed=(mesh_change is not None and pml_change is not None
        and mesh_change<=tolerance_percent and pml_change<=tolerance_percent)
    return dict(base=base,mesh=mesh,pml=pml,
                mesh_last_two_change_pct=mesh_change,
                pml_last_two_change_pct=pml_change,
                tolerance_percent=tolerance_percent,
                last_two_sweeps_passed=passed,
                is_fully_validated=False,
                caveat="Even a pass requires domain-size, polarization, mode tracking, cladding and experimental convergence.")

if __name__=="__main__":
    import argparse,json,time
    parser=argparse.ArgumentParser(description=__doc__)
    for name,field in Fiber.__dataclass_fields__.items():
        val=field.default
        parser.add_argument("--"+name.replace("_","-"),type=type(val),default=val)
    parser.add_argument("--convergence",action="store_true",
                        help="repeat independent runs on 3 grids and 3 PML thicknesses")
    args=vars(parser.parse_args())
    isstudy=args.pop("convergence")
    f=Fiber(**args)
    start=__import__("time").time()
    result=convergence_study(f) if isstudy else solve(f)
    result["elapsed_s"]=__import__("time").time()-start
    print(json.dumps(result,indent=2))
