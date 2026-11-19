"""Equivalent sparse pencils for the pre-existing mixed quadratic-beta system.

For Et=e, Ez=i*beta*phi:
  A=[-Ktt,0;0,0], B=[Mtt,i*Ktz;-i*Kzt,Kzz], A x=beta² B x.
The original companion is retained as an independent linearization check.
"""
import numpy as np
from scipy.sparse import bmat, csc_matrix, eye
from scipy.sparse.linalg import eigs, splu, LinearOperator


def spectrum(system,target,n_modes,options):
    K0,K1,K2=(system["K"+str(i)] for i in range(3));n=K0.shape[0]
    backend=options.get("backend","beta_squared")
    limit=options.get("max_dofs",80000)
    if n>limit:raise MemoryError(f"mixed DOFs {n} exceed explicit solver.max_dofs={limit}; refine locally or use HPC")
    sigma=system["k0"]*target
    if backend=="companion":
        I=eye(n,format="csc",dtype=complex);Z=csc_matrix((n,n),dtype=complex)
        A=bmat([[Z,I],[-K0,-K1]],format="csc");B=bmat([[I,Z],[Z,K2]],format="csc")
        shift=sigma
    elif backend=="beta_squared":
        nt=int(np.count_nonzero(system["keep"]<system["ne"]))
        zero=csc_matrix((n-nt,n-nt),dtype=complex)
        A=bmat([[-K0[:nt,:nt],None],[None,zero]],format="csc")
        B=bmat([[K2[:nt,:nt],1j*K1[:nt,nt:]],[-1j*K1[nt:,:nt],K0[nt:,nt:]]],format="csc")
        shift=sigma*sigma
    else:raise ValueError("solver.backend must be beta_squared or companion")
    real_matrix=(not np.any(A.data.imag) and not np.any(B.data.imag))
    if real_matrix:A=A.real.tocsc();B=B.real.tocsc()
    factor=splu(A-shift*B)
    op=LinearOperator(A.shape,matvec=lambda x:factor.solve(B@x),dtype=A.dtype)
    k=min(max(2*n_modes+8,12),A.shape[0]-2)
    rng=np.random.default_rng(int(options.get("seed",42)))
    initial=rng.standard_normal(A.shape[0])
    if not real_matrix:initial=initial+1j*rng.standard_normal(A.shape[0])
    mu,vec=eigs(op,k=k,which="LM",tol=options.get("tolerance",1e-10),maxiter=int(options.get("max_iterations",4000)),
                ncv=min(A.shape[0],max(2*k+1,40)),v0=initial)
    candidates=[]
    for j,value in enumerate(mu):
        if abs(value)<1e-14 or not np.isfinite(value):continue
        eigenvalue=shift+1/value
        beta=eigenvalue if backend=="companion" else np.sqrt(eigenvalue+0j)
        if beta.real<0:beta=-beta
        v=vec[:,j]
        residual=float(np.linalg.norm(A@v-eigenvalue*(B@v))/(np.linalg.norm(A@v)+abs(eigenvalue)*np.linalg.norm(B@v)+1e-30))
        if backend=="companion":e=v[:n].copy()
        else:e=np.concatenate((v[:nt],1j*beta*v[nt:]))
        candidates.append((beta,e,residual))
    return candidates,{"backend":backend,"mixed_dofs":n,"pencil_dofs":A.shape[0],
                       "real_pencil":bool(real_matrix),"target_neff":target,"shift":shift,
                       "tolerance":options.get("tolerance",1e-10),"seed":options.get("seed",42)}
