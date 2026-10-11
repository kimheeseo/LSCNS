"""M8 symmetric two-core weak-guidance scalar FEM supermodes and independent
hollow capillary analytical baseline. NOT antiresonant HCF/NANF FEM or loss.
"""
from pathlib import Path
import sys,numpy as np,json
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/'waveoptics_m7'/'src'))
from fiber_modes import Fiber,solve_modes,C0
U01=2.404825557695773
def solve_two_core(separation_um=16.,n=101,halfwidth_um=27.):
 if separation_um<=8.5 or halfwidth_um<=separation_um/2+9:raise ValueError("invalid 2-core cross section")
 f=Fiber(domain_halfwidth_um=halfwidth_um)
 sol=solve_modes(f,n,cores=[(-separation_um/2,0),(separation_um/2,0)],modes=2)
 if not all(sol['guided']):raise RuntimeError("two modes not guided")
 even,odd=map(float,sol['neff']);split=even-odd
 parity=[]
 for k in range(2):
  q=sol['fields'][:,k].reshape((n,n));norm=np.linalg.norm(q)
  parity.append(dict(even_score=float(np.linalg.norm(q-q[:,::-1])/norm),odd_score=float(np.linalg.norm(q+q[:,::-1])/norm)))
 return dict(separation_um=float(separation_um),n_grid=n,n_even=even,n_odd=odd,index_split=split,
    coupling_length_mm=float(f.wavelength_um/(2*split)/1000),parity=parity,residuals=sol['residuals'],solution=sol)
def capillary_neff(w,radius_um=15.,n_air=1.00027):
 wl=np.asarray(w,float)
 if radius_um<=0 or n_air<1 or np.any(wl<=0):raise ValueError("invalid capillary")
 z=n_air*n_air-(U01*wl/(2*np.pi*radius_um))**2
 if np.any(z<=0):raise ValueError("outside capillary propagation regime")
 return np.sqrt(z)
def capillary_dispersion(w=1.55):
 h=.002;n=lambda x:float(capillary_neff(x))
 d2=(-n(w+2*h)+16*n(w+h)-30*n(w)+16*n(w-h)-n(w-2*h))/(12*h*h)
 return float(-(w/C0)*1e12*d2)
def run(output_dir):
 import matplotlib
 matplotlib.use('Agg')
 import matplotlib.pyplot as plt
 out=Path(output_dir);out.mkdir(parents=True,exist_ok=True)
 ss=[12.,16.,20.,24.];r=[solve_two_core(x) for x in ss]
 for a in r:
  assert a['index_split']>0 and max(a['residuals'])<1e-6
  assert a['parity'][0]['even_score']<.04 and a['parity'][1]['odd_score']<.04
 assert all(r[i]['index_split']>r[i+1]['index_split'] for i in range(len(r)-1))
 wl=np.linspace(1.3,1.7,101);hcf=capillary_neff(wl)
 summary=dict(MCF=[{k:v for k,v in z.items() if k!='solution'} for z in r],
              HCF_capillary=dict(neff_1p55=float(capillary_neff(1.55)),D_1p55_ps_per_nm_km=capillary_dispersion(),
                                 n_air=1.00027,radius_um=15.,model='ideal capillary analytic; no antiresonant leakage or PML'))
 (out/'m8_summary.json').write_text(json.dumps(summary,indent=2))
 sol=r[1]['solution'];n=sol['n'];XY=sol['xy'];fields=sol['fields'];X=XY[:,0].reshape(n,n);Y=XY[:,1].reshape(n,n)
 fig,ax=plt.subplots(1,2,figsize=(10,4),constrained_layout=True)
 for k,title in enumerate(('Even supermode','Odd supermode')):
  q=fields[:,k].reshape(n,n)
  im=ax[k].pcolormesh(X,Y,q,shading='auto',cmap='RdBu_r',vmin=-abs(q).max(),vmax=abs(q).max())
  ax[k].set(xlabel='x (um)',ylabel='y (um)',title=title,aspect='equal');fig.colorbar(im,ax=ax[k])
 fig.savefig(out/'m8_mcf_supermodes.png',dpi=140);plt.close(fig)
 fig,ax=plt.subplots(1,2,figsize=(10,4),constrained_layout=True)
 ax[0].semilogy(ss,[z['index_split'] for z in r],'o-');ax[0].set(xlabel='core pitch (um)',ylabel='delta n_eff',title='MCF supermode coupling');ax[0].grid(alpha=.3)
 ax[1].plot(wl,hcf);ax[1].set(xlabel='wavelength (um)',ylabel='n_eff',title='Ideal hollow capillary reference');ax[1].ticklabel_format(axis='y',useOffset=False);ax[1].grid(alpha=.3)
 fig.savefig(out/'m8_coupling_and_hcf.png',dpi=140);plt.close(fig)
 return summary
