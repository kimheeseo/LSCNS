"""Independent real-space beta=0 square TE/TM finite-difference reference.

TE Hz: -div(eps^-1 grad Hz)=k0² Hz, harmonic face coefficient.
TM Ez: -Delta Ez=k0² eps Ez. Bloch phases at boundaries.
This is a separate discretization from vector spectral PWE, not experimental data.
"""
import copy
import numpy as np
from scipy.sparse import coo_matrix
from scipy.sparse.linalg import eigsh
from step5_bloch import assemble_operator,dielectric_grid

def reference_te_tm(config,k_frac,grid_n=64):
 if config['lattice']!='square':raise ValueError('independent reference is square, beta=0 only')
 c=copy.deepcopy(config);c.update(schema='LSCNS_STEP5_BLOCH_INPUT_V083',grid_n=grid_n,beta_norms=[.001]);c.pop('backend',None);c.pop('pwe_cutoff',None)
 K,M=assemble_operator(c,k_frac,0.)
 tm=eigsh(K,k=config['bands'],M=M,sigma=-1e-8,which='LM',return_eigenvectors=False)
 eps,_=dielectric_grid(c);inv=1/eps;N=grid_n;h=config['pitch_um']/N;rr=[];cc=[];vv=[]
 for i in range(N):
  for j in range(N):
   row=i*N+j;diag=0.
   for du,dv in ((1,0),(-1,0),(0,1),(0,-1)):
    ii=i+du;jj=j+dv;col=(ii%N)*N+jj%N;a=2*inv[row]*inv[col]/(inv[row]+inv[col])/h**2;diag+=a
    phase=np.exp(2j*np.pi*(k_frac[0]*(ii//N)+k_frac[1]*(jj//N)));rr.append(row);cc.append(col);vv.append(-a*phase)
   rr.append(row);cc.append(row);vv.append(diag)
 T=coo_matrix((vv,(rr,cc)),shape=(N*N,N*N)).tocsc()
 te=eigsh(T,k=config['bands'],sigma=-1e-8,which='LM',return_eigenvectors=False)
 return np.sort(np.sqrt(np.maximum(np.r_[te.real,tm.real],0))*config['pitch_um']/(2*np.pi))[:config['bands']]
