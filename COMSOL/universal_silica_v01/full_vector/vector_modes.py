"""Research-grade prototype 2-D full-vector finite-difference eigenmode solver.
Isotropic nonmagnetic z-invariant dielectrics, lossless; no PML.
The Ht=(Hx,Hy) vectorial operator is derived from curl epsilon^-1 curl H=k0² H
and div(H)=0; Ez is reconstructed from Maxwell curl H.
All lengths in um, eps_r is relative dielectric function.
"""
from dataclasses import dataclass,asdict
from pathlib import Path
import json
import numpy as np
from scipy.sparse import diags,eye,kron,bmat
from scipy.sparse.linalg import eigs
from scipy.special import jv,kv,jvp,kvp
from scipy.optimize import brentq

B=(0.6961663,0.4079426,0.8974794)
C=(0.0684043**2,0.1162414**2,9.896161**2)
C_LIGHT=299792458.0

def silica_index(w_um):
    if not .21<w_um<3.7: raise ValueError('wavelength outside silica Sellmeier domain')
    return float(np.sqrt(1+sum(b*w_um*w_um/(w_um*w_um-c) for b,c in zip(B,C))))

@dataclass(frozen=True)
class Configuration:
    wavelength_nm:float=1550.
    profile:str='step'
    core_radius_um:float=5.5
    delta_n:float=.0034785793846155
    window_halfwidth_um:float=22.
    grid_n:int=77
    graded_alpha:float=2.
    trench_start_um:float=7.
    trench_end_um:float=10.
    trench_delta_n:float=-.001
    shells:tuple=((5.,.004),(7.,0.),(9.,-.0007))
    def validate(self):
        if self.profile not in ('step','graded','trench','multilayer'):raise ValueError('unsupported profile')
        for n in ('wavelength_nm','core_radius_um','window_halfwidth_um'):
            if not np.isfinite(getattr(self,n)) or getattr(self,n)<=0:raise ValueError(n)
        if not(250<self.wavelength_nm<3600):raise ValueError('wavelength')
        if not (self.window_halfwidth_um>2.1*self.core_radius_um):raise ValueError('domain too small')
        if not isinstance(self.grid_n,int) or self.grid_n<29 or self.grid_n>181 or self.grid_n%2==0:raise ValueError('grid_n odd [29,181]')
        if not np.isfinite(self.delta_n) or self.delta_n<=0 or self.delta_n>.1:raise ValueError('delta_n')
        if self.profile=='trench' and not self.core_radius_um<self.trench_start_um<self.trench_end_um<self.window_halfwidth_um:raise ValueError('trench')
        if self.profile=='multilayer':
            if not self.shells or self.shells[0][0]<=0 or self.shells[-1][0]>=self.window_halfwidth_um:raise ValueError('shells')
            if not all(self.shells[i][0]<self.shells[i+1][0] for i in range(len(self.shells)-1)):raise ValueError('shell boundaries')
        return self

def make_index(c:Configuration):
    c.validate()
    n=c.grid_n;L=c.window_halfwidth_um
    xy=np.linspace(-L,L,n);dx=xy[1]-xy[0]
    xx,yy=np.meshgrid(xy,xy,indexing='xy');r=np.hypot(xx,yy)
    nc=silica_index(c.wavelength_nm/1000)
    if c.profile=='step':offset=np.where(r<c.core_radius_um,c.delta_n,0.)
    elif c.profile=='graded':offset=np.where(r<c.core_radius_um,c.delta_n*np.maximum(1-(r/c.core_radius_um)**c.graded_alpha,0.),0.)
    elif c.profile=='trench':offset=np.where(r<c.core_radius_um,c.delta_n,np.where((r>=c.trench_start_um)&(r<c.trench_end_um),c.trench_delta_n,0.))
    else:
        offset=np.zeros_like(r);prev=0.
        for end,delta in c.shells:
            offset=np.where((r>=prev)&(r<end),delta,offset);prev=end
    return xy,nc+offset

def operators(n:int,d:float):
    k=n-2
    D2=diags([np.ones(k-1),-2*np.ones(k),np.ones(k-1)],[-1,0,1],shape=(k,k),format='csr')/d**2
    D1=diags([-np.ones(k-1),np.ones(k-1)],[-1,1],shape=(k,k),format='csr')/(2*d)
    eye1=eye(k,format='csr')
    Dx=kron(eye1,D1,format='csr');Dy=kron(D1,eye1,format='csr')
    Lap=kron(eye1,D2,format='csr')+kron(D2,eye1,format='csr')
    return Dx,Dy,Lap

def integrate_intensity(I,xy):
    """Full transverse-field intensity integrals over dxdy (finite-difference quadrature).

    I = |Ex|^2 + |Ey|^2 + |Ez|^2, so Aeff = (integral I)^2/(integral I^2).
    The moment MFD is 2 sqrt(2 <r^2>) centered on the intensity centroid.
    A separate Petermann-II-like near-field estimate uses f = sqrt(I).
    """
    I=np.asarray(I,dtype=float);xy=np.asarray(xy,dtype=float)
    if I.shape!=(len(xy),len(xy)) or np.any(I<0) or not np.isfinite(I).all():raise ValueError('intensity grid')
    d=xy[1]-xy[0];A=d*d;power=float(np.sum(I)*A)
    if power<=0:raise ValueError('zero intensity')
    xx,yy=np.meshgrid(xy,xy)
    cx=float(np.sum(I*xx)*A/power);cy=float(np.sum(I*yy)*A/power)
    moment=float(np.sum(((xx-cx)**2+(yy-cy)**2)*I)*A/power)
    aeff=power**2/(float(np.sum(I*I))*A)
    f=np.sqrt(I);fy,fx=np.gradient(f,d,d,edge_order=2)
    gradient=float(np.sum(fx*fx+fy*fy)*A)
    return dict(Aeff_um2=aeff,MFD_2nd_um=2*np.sqrt(2*moment),
                MFD_PII_um=2*np.sqrt(2*power/gradient) if gradient>0 else float('nan'),
                centroid_x_um=cx,centroid_y_um=cy,power_integral=power)


def solve(c:Configuration,modes=3):
    xy,ni=make_index(c);n=c.grid_n;d=xy[1]-xy[0];k0=2*np.pi/(c.wavelength_nm/1000)
    eps=ni**2;gy,gx=np.gradient(np.log(eps),d,d,edge_order=2)
    select=np.s_[1:-1,1:-1];e=eps[select].ravel();p=ni[select].ravel()
    gx=diags(gx[select].ravel(),format='csr');gy=diags(gy[select].ravel(),format='csr')
    Dx,Dy,Lap=operators(n,d)
    base=Lap+diags(k0*k0*e,format='csr')
    # H field PDE from Maxwell's curl-eps-curl eigenmode and div H=0:
    # beta² Hx = (lap + k0² eps) Hx - (d_y ln eps) (d_y Hx - d_x Hy)
    # beta² Hy = (lap + k0² eps) Hy + (d_x ln eps) (d_y Hx - d_x Hy)
    M=bmat([[base-gy@Dy,gy@Dx],[gx@Dy,base-gx@Dx]],format='csc')
    v=np.zeros(2*len(e));v[len(e)//2]=1;v[len(e)+len(e)//2]=.7
    number=min(modes,6)
    vals,vecs=eigs(M,k=number,sigma=(k0*max(e)**.5)**2+.005,which='LM',v0=v,tol=2e-9,maxiter=3000)
    modes_out=[]
    for j in np.argsort(vals.real)[::-1]:
        eig=vals[j];beta=np.sqrt(eig);neff=beta/k0
        if abs(neff.imag)>2e-4:continue
        if not (ni[0,0]+1e-6 < neff.real < ni.max()+2e-3):continue
        Hx=np.zeros((n,n),complex);Hy=np.zeros((n,n),complex)
        Hx[select]=vecs[:len(e),j].reshape(n-2,n-2);Hy[select]=vecs[len(e):,j].reshape(n-2,n-2)
        divH=np.gradient(Hx,d,axis=1)+np.gradient(Hy,d,axis=0)
        Hz=1j*divH/beta
        curlx=np.gradient(Hz,d,axis=0)-1j*beta*Hy
        curly=1j*beta*Hx-np.gradient(Hz,d,axis=1)
        curlz=np.gradient(Hy,d,axis=1)-np.gradient(Hx,d,axis=0)
        Ex=1j*curlx/(k0*eps);Ey=1j*curly/(k0*eps);Ez=1j*curlz/(k0*eps)
        I=np.abs(Ex)**2+np.abs(Ey)**2+np.abs(Ez)**2
        metrics=integrate_intensity(I,xy)
        xx,yy=np.meshgrid(xy,xy)
        hres=np.linalg.norm(M@vecs[:,j]-eig*vecs[:,j])/np.linalg.norm(eig*vecs[:,j])
        px=float(np.sum(np.abs(Ex)**2)/np.sum(I));py=float(np.sum(np.abs(Ey)**2)/np.sum(I));pz=float(np.sum(np.abs(Ez)**2)/np.sum(I))
        corepower=I[np.hypot(xx,yy)<=c.core_radius_um].sum()/I.sum()
        edgepower=(I[:3,:].sum()+I[-3:,:].sum()+I[:,:3].sum()+I[:,-3:].sum())/I.sum()
        modes_out.append(dict(neff=float(neff.real),neff_imag_numerical=float(neff.imag),Aeff_um2=float(metrics["Aeff_um2"]),MFD_2nd_um=float(metrics["MFD_2nd_um"]),MFD_PII_um=float(metrics["MFD_PII_um"]),
            residual=float(hres),Ex_intensity_fraction=px,Ey_intensity_fraction=py,Ez_intensity_fraction=pz,core_power_fraction=float(corepower),edge_power_fraction=float(edgepower),field_intensity=I/I.max(),
            Ex=Ex,Ey=Ey,Ez=Ez,Hx=Hx,Hy=Hy,Hz=Hz))
    if not modes_out:raise RuntimeError('no guided vector mode found; adjust grid/domain')
    # sort by field symmetry / highest neff pair; both polarization partners may differ due to mesh
    return dict(config=c,xy_um=xy,index_map=ni,modes=modes_out,method='2D transverse Hx Hy full-vector finite difference; Ez reconstructed from Maxwell')


def he11_exact_step(c:Configuration):
    """Independent exact full-vector HE11 characteristic equation, circular step-index only."""
    if c.profile!='step':raise ValueError('HE11 oracle step only')
    lam=c.wavelength_nm/1000;k0=2*np.pi/lam;nc=silica_index(lam);nco=nc+c.delta_n;a=c.core_radius_um
    v=k0*a*np.sqrt(nco*nco-nc*nc)
    if not 0<v<2.4048255:raise ValueError('comparison restricted to V<2.405')
    def f(u):
        w=np.sqrt(v*v-u*u)
        A=jvp(1,u)/(u*jv(1,u));B=kvp(1,w)/(w*kv(1,w))
        ne2=nco*nco-u*u/(k0*a)**2
        return (A+B)*(nco*nco*A+nc*nc*B)-ne2*(1/u/u+1/w/w)**2
    us=np.linspace(.001,v-.001,300);zs=[f(u) for u in us]
    roots=[brentq(f,a,b) for a,b,fa,fb in zip(us[:-1],us[1:],zs[:-1],zs[1:]) if np.isfinite(fa) and np.isfinite(fb) and fa*fb<0]
    if not roots:raise RuntimeError('could not bracket exact HE11 root')
    return float(np.sqrt(nco*nco-(min(roots)/(k0*a))**2))


def export_result(sol,path,field_stride=2):
    """Portable JSON for browser import. Full complex fields remain available in .npz."""
    c=sol['config']; xy=sol['xy_um']; stride=int(field_stride)
    if stride<1:raise ValueError('field_stride >= 1')
    out={"schema":"LSCNS_vector_mode_v1", "method":sol['method'],
         "input":asdict(c),"modes":[],
         "units":{"neff":"dimensionless","MFD_2nd_um":"um","MFD_PII_um":"um","Aeff_um2":"um^2"},
         "validation":{"domain":"transverse square, zero-field Dirichlet exterior; no radiation PML",
         "scope":"full-vector *finite difference*, NOT FEM; isotropic z-invariant silica profiles",
         "limitation":"Preliminary; staircase interfaces and finite window can shift neff; not a measured product or certified loss result"}}
    for idx,m in enumerate(sol['modes']):
        base={k:v for k,v in m.items() if isinstance(v,float)}
        base['mode_index']=idx+1
        if idx==0:
            norm=max(float(m['field_intensity'].max()),1e-25)
            fields={key:(np.abs(m[key])**2/norm)[::stride,::stride] for key in ('Ex','Ey','Ez')}
            fields['Itotal']=m['field_intensity'][::stride,::stride]
            base['field_map']={'x_um':np.round(xy[::stride],5).tolist(),
                'grid_size':len(xy[::stride]),
                'intensity_normalized_to_max_total':{key:np.round(value,7).tolist() for key,value in fields.items()},
                'index_map':np.round(sol['index_map'][::stride,::stride],7).tolist()}
        out['modes'].append(base)
    Path(path).write_text(json.dumps(out,ensure_ascii=False,indent=2))
    return out


def save_full_fields(sol,path):
    """Save true complex E/H field components for reproducibility and later research."""
    m=sol['modes'][0]
    np.savez_compressed(path,x_um=sol['xy_um'],index=sol['index_map'],
        Ex=m['Ex'],Ey=m['Ey'],Ez=m['Ez'],Hx=m['Hx'],Hy=m['Hy'],Hz=m['Hz'],
        I=m['field_intensity'])

if __name__=='__main__':
    import sys
    c=Configuration(grid_n=int(sys.argv[1]) if len(sys.argv)>1 else 57)
    r=solve(c,modes=4)
    print('reference',he11_exact_step(c))
    for m in r['modes']:
        print({k:round(v,10) for k,v in m.items() if isinstance(v,float)})
    export_result(r,Path(__file__).parent/'vector_mode_result.json')
