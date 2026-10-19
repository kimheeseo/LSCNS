"""Independent radial P1 eigenproblem for the ideal scalar hollow capillary.

-(r psi')'=kt^2*r*psi, regular at r=0, psi(R)=0.
This verifies the J0/Dirichlet baseline only, not ARF/NANF glass-wall physics.
"""
import numpy as np
from scipy.sparse import coo_matrix
from scipy.sparse.linalg import eigsh
from numpy.polynomial.legendre import leggauss
C0=299792458.


def solve_capillary(wavelength_um=1.55,radius_um=15.,n_air=1.00027,elements=320,return_field=False):
    if not isinstance(elements,(int,np.integer)) or isinstance(elements,bool) or elements<4:
        raise ValueError('elements must be integer >=4')
    if not np.all(np.isfinite([wavelength_um,radius_um,n_air])) or wavelength_um<=0 or radius_um<=0 or n_air<1:
        raise ValueError('invalid capillary parameters')
    r=np.linspace(0,radius_um,elements+1);h=np.diff(r)
    x,w=leggauss(4);t=(x+1)/2;phi=np.stack([1-t,t],1)
    rq=r[:-1,None]+h[:,None]*t;weights=h[:,None]*w/2
    mass=np.einsum('eq,qi,qj->eij',weights*rq,phi,phi)
    stiff=((r[1:]**2-r[:-1]**2)/(2*h**2))[:,None,None]*np.array([[1.,-1.],[-1.,1.]])
    edges=np.stack([np.arange(elements),np.arange(elements)+1],1)
    rows=np.broadcast_to(edges[:,:,None],mass.shape).ravel();cols=np.broadcast_to(edges[:,None,:],mass.shape).ravel()
    def mat(v):return coo_matrix((v.ravel(),(rows,cols)),shape=(elements+1,elements+1)).tocsr()[:-1,:-1]
    K,M=mat(stiff),mat(mass)
    vals,vec=eigsh(K,k=1,M=M,sigma=0,which='LM',tol=1e-12,v0=np.ones(elements))
    kt2=float(vals[0]);v=vec[:,0];field=np.r_[v,0.]
    if field[0]<0:field=-field
    q=kt2/(2*np.pi)**2;z=n_air*n_air-q*wavelength_um**2
    if z<=0:raise ValueError('outside capillary propagation regime')
    ne=np.sqrt(z);D=wavelength_um/C0*1e12*q*n_air**2/ne**3
    amp=field[edges]@phi.T;wt=2*np.pi*weights*rq;power=np.sum(wt*amp**2)
    result=dict(elements=elements,radial_step_um=radius_um/elements,wavelength_um=wavelength_um,
                radius_um=radius_um,n_air=n_air,transverse_kt2_per_um2=kt2,
                n_eff=float(ne),D_ps_nm_km=float(D),Aeff_um2=float(power**2/np.sum(wt*amp**4)),
                residual=float(np.linalg.norm(K@v-kt2*(M@v))/np.linalg.norm(kt2*(M@v))))
    if return_field:result.update(radius=r,field=field/np.sqrt(power))
    return result
