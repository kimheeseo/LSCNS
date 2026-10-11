"""M4 reproducible 3-D PEC-cavity baseline: lowest-order tetrahedral Nedelec FEM.
Length in micrometres; relative isotropic eps=n**2; unit-free eigenvalue (k0*um)^2.
This is a VERIFIED CAVITY EIGENMODE subset, NOT a restored 3D PML/port solver.
"""
from __future__ import annotations
import json
from pathlib import Path
import numpy as np
from scipy.sparse import coo_matrix
from scipy.sparse.linalg import eigsh
C0=299792458.0
LOCAL_EDGES=[(0,1),(0,2),(0,3),(1,2),(1,3),(2,3)]
SPLITS=[(0,1,3,7),(0,3,2,7),(0,2,6,7),(0,6,4,7),(0,4,5,7),(0,5,1,7)]
def tetra_grid(lengths=(2.,1.,3.),shape=(4,3,6)):
    if len(lengths)!=3 or len(shape)!=3 or min(lengths)<=0 or min(shape)<2:raise ValueError("positive lengths and shape>=2")
    nx,ny,nz=map(int,shape);Lx,Ly,Lz=lengths
    xyz=np.array([[Lx*i/nx,Ly*j/ny,Lz*k/nz] for k in range(nz+1) for j in range(ny+1) for i in range(nx+1)],float)
    def node(i,j,k):return k*(ny+1)*(nx+1)+j*(nx+1)+i
    cells=[]
    for k in range(nz):
      for j in range(ny):
       for i in range(nx):
        cub=[node(i,j,k),node(i+1,j,k),node(i,j+1,k),node(i+1,j+1,k),node(i,j,k+1),node(i+1,j,k+1),node(i,j+1,k+1),node(i+1,j+1,k+1)]
        cells.extend([[cub[q] for q in tet] for tet in SPLITS])
    return xyz,np.asarray(cells,dtype=int)
def assemble(lengths=(2.,1.,3.),shape=(4,3,6),index=1.):
    if not np.isfinite(index) or index<=0:raise ValueError("index must be positive")
    xyz,tets=tetra_grid(lengths,shape)
    all_edges=sorted({tuple(sorted((int(t[a]),int(t[b])))) for t in tets for a,b in LOCAL_EDGES})
    edge_idx={e:i for i,e in enumerate(all_edges)}
    nn=len(all_edges);rows=[];cols=[];kv=[];mv=[]
    # Degree-two exact four-point tetra integration.
    qA=.5854101966249685;qB=.1381966011250105
    quad=np.full((4,4),qB);np.fill_diagonal(quad,qA)
    for t in tets:
      pts=xyz[t];B=(pts[1:]-pts[0]).T;det=np.linalg.det(B);volume=abs(det)/6.
      if volume<1e-18:raise ValueError("degenerate tetra")
      grads=np.empty((4,3));grads[1:]=np.linalg.inv(B);grads[0]=-grads[1:].sum(axis=0)
      basis=np.asarray([np.asarray([q[i]*grads[j]-q[j]*grads[i] for i,j in LOCAL_EDGES]) for q in quad])
      curls=np.asarray([2*np.cross(grads[i],grads[j]) for i,j in LOCAL_EDGES])
      localK=volume*(curls@curls.T)
      localM=volume*index**2*np.einsum('qic,qjc->ij',basis,basis)/4.
      ids=[];sign=[]
      for i,j in LOCAL_EDGES:
       a,b=int(t[i]),int(t[j]);ids.append(edge_idx[min(a,b),max(a,b)]);sign.append(1 if a<b else -1)
      sg=np.asarray(sign)
      localK*=sg[:,None]*sg[None,:];localM*=sg[:,None]*sg[None,:]
      rr=np.repeat(ids,6);cc=np.tile(ids,6)
      rows.extend(rr);cols.extend(cc);kv.extend(localK.ravel());mv.extend(localM.ravel())
    K=coo_matrix((kv,(rows,cols)),shape=(nn,nn)).tocsr()
    M=coo_matrix((mv,(rows,cols)),shape=(nn,nn)).tocsr()
    L=np.asarray(lengths)
    fixed=[]
    for k,(i,j) in enumerate(all_edges):
      a,b=xyz[i],xyz[j]
      if any(np.isclose(a[q],0) and np.isclose(b[q],0) or np.isclose(a[q],L[q]) and np.isclose(b[q],L[q]) for q in range(3)):fixed.append(k)
    free=np.setdiff1d(np.arange(nn),fixed)
    return xyz,tets,all_edges,free,K,M
def analytic_te101(lengths=(2.,1.,3.),index=1.):
    x,_,z=lengths
    return C0/(2e-6*index)*np.sqrt(1/x**2+1/z**2)
def solve(lengths=(2.,1.,3.),shape=(4,3,6),index=1.):
    xyz,tets,edges,free,K,M=assemble(lengths,shape,index)
    lam=(2*np.pi*analytic_te101(lengths,index)*1e-6/C0)**2
    vals,vecs=eigsh(K[free][:,free],k=min(8,len(free)-2),M=M[free][:,free],sigma=lam*.95,which='LM',tol=1e-8)
    pos=np.flatnonzero((vals>lam*.25)&np.isfinite(vals))
    if not len(pos):raise RuntimeError("No positive cavity mode")
    order=pos[np.argsort(vals[pos])]
    eigen=float(vals[order[0]])
    mode=np.zeros(K.shape[0]);mode[free]=vecs[:,order[0]]
    f=C0*np.sqrt(eigen)/(2*np.pi*1e-6)
    exact=analytic_te101(lengths,index)
    rel=abs(f/exact-1)
    res=np.linalg.norm(K[free][:,free]@mode[free]-eigen*M[free][:,free]@mode[free])/max(np.linalg.norm(eigen*M[free][:,free]@mode[free]),1e-20)
    return {"lengths_um":list(lengths),"shape":list(shape),"index":float(index),"nodes":int(len(xyz)),"tetrahedra":int(len(tets)),"edges":int(len(edges)),"free_dofs":int(len(free)),"frequency_hz":float(f),"analytic_te101_hz":float(exact),"relative_error":float(rel),"eigenvalue":eigen,"relative_residual":float(res)},(xyz,tets,edges,mode)
def save_plot(mesh,result,out):
    import matplotlib;matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    xyz,tets,edges,mode=mesh;cent=np.array([(xyz[a]+xyz[b])/2 for a,b in edges])
    lengths=result["lengths_um"]
    select=np.abs(cent[:,1]-lengths[1]/2)<lengths[1]*.27
    fig=plt.figure(figsize=(11,4.5),constrained_layout=True);ax=fig.add_subplot(121,projection="3d")
    ax.scatter(cent[select,0],cent[select,1],cent[select,2],c=np.abs(mode[select]),s=3,cmap="viridis",alpha=.85)
    ax.set(xlabel="x (um)",ylabel="y (um)",zlabel="z (um)",title="Nedelec edge field amplitude (PEC cavity)")
    ax2=fig.add_subplot(122);pts=cent[select];b=ax2.scatter(pts[:,0],pts[:,2],c=np.abs(mode[select]),s=10,cmap="viridis")
    ax2.set(xlabel="x (um)",ylabel="z (um)",title="Approximate central-section edge coefficients");fig.colorbar(b,ax=ax2,label="|edge DOF| (arbitrary)")
    fig.suptitle("M4 reconstructed verification baseline — 3D FEM cavity, not PML/port S")
    fig.savefig(out,dpi=140);plt.close(fig)
def run(outdir):
    outdir=Path(outdir);outdir.mkdir(parents=True,exist_ok=True)
    rec=[];finemesh=None
    for shape in [(3,2,4),(4,3,6),(5,4,8)]:
      info,mesh=solve(shape=shape);rec.append(info);finemesh=mesh
    (outdir/"m4_cavity_results.json").write_text(json.dumps(rec,indent=2))
    save_plot(finemesh,rec[-1],outdir/"m4_cavity_field.png")
    import matplotlib;matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.figure(figsize=(6,4));plt.plot([r["tetrahedra"] for r in rec],[r["relative_error"] for r in rec],"o-");plt.yscale("log")
    plt.xlabel("Tetrahedra");plt.ylabel("Relative frequency error");plt.title("M4 3D PEC cavity convergence (TE101)")
    plt.grid(True,alpha=.3);plt.tight_layout();plt.savefig(outdir/"m4_convergence.png",dpi=160);plt.close()
    return rec
if __name__=="__main__":
 import argparse
 p=argparse.ArgumentParser();p.add_argument("--out",default="results");args=p.parse_args()
 print(json.dumps(run(args.out),indent=2))
