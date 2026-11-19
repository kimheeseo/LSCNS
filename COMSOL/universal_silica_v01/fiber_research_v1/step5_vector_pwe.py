"""STEP5 v0.8.7 vector Maxwell Bloch PWE; formulas/units fixed here.

H=sum_G U_G h_G exp(i (k+G).r+i beta z-i omega t), q_G.H_G=0.
M_ia,jb=eta_(Gi-Gj) (q_i cross U_ia)*.(q_j cross U_jb), eta=1/eps.
M h=(omega/c)^2 h. Use pitch=1 internally; f=omega*pitch/(2*pi*c)=pitch/lambda.
beta_norm=beta*pitch/(2*pi); light lines beta/k0=n_air,n_host.
Circular eta_G=eta_host delta_G0+(eta_air-eta_host)*pi*r²/area*
  2*J1(|G|r)/(|G|r)*sum_sites exp(-i G.site). No raster epsilon approximation.
Real nondispersive isotropic media. Finite cutoff/k mesh gaps are candidates only.
Source theory: https://mpb.readthedocs.io/en/stable/Developer_Information/
This module implements independent small dense PWE, NOT the MPB executable.
"""
import copy,hashlib,json,math
import numpy as np
from scipy.linalg import eigh
from scipy.special import j1
from step5_bloch import primitive,MOTIFS,smallest_separation,high_symmetry_path,_kmesh_points,hole_mask
INPUT='LSCNS_STEP5_BLOCH_INPUT_V087';OUTPUT='LSCNS_STEP5_BLOCH_RESULT_V087'

def example_config():
 return {'schema':INPUT,'schema_version':'1.0.0','backend':'vector_pwe','lattice':'triangular','pitch_um':1.5,'hole_diameter_over_pitch':.9,
         'n_host':1.45,'n_air':1.,'grid_n':32,'bands':12,'path_points_per_segment':4,'kmesh_n':3,'beta_norms':[.6,1.,1.4],
         'pwe_cutoff':3,'input_basis':'explicit illustrative constant indices, NOT a commercial silica/gas specification'}

def validate(c):
 if c.get('schema')!=INPUT or c.get('schema_version')!='1.0.0' or c.get('backend')!='vector_pwe':raise ValueError('vector PWE input schema/backend required')
 for key,lo,hi in (('pitch_um',.1,1000),('hole_diameter_over_pitch',.01,.95),('n_host',.9,5),('n_air',.9,5)):
  x=c.get(key)
  if isinstance(x,bool) or not isinstance(x,(int,float)) or not np.isfinite(x) or not lo<=x<=hi:raise ValueError('invalid '+key)
 if c['lattice'] not in MOTIFS:raise ValueError('unknown lattice')
 if c['hole_diameter_over_pitch']*c['pitch_um']>=.98*smallest_separation(c['lattice'],c['pitch_um']):raise ValueError('overlapping circular motif')
 for key,lo,hi in (('pwe_cutoff',1,8),('bands',2,24),('grid_n',10,96),('path_points_per_segment',2,30),('kmesh_n',2,12)):
  if type(c.get(key))!=int or not lo<=c[key]<=hi:raise ValueError('invalid integer '+key)
 if c['bands']>=2*(2*c['pwe_cutoff']+1)**2:raise ValueError('too many bands')
 b=c.get('beta_norms')
 if not isinstance(b,list) or not 1<=len(b)<=10 or any(isinstance(x,bool) or not isinstance(x,(int,float)) or not np.isfinite(x) or not 0<=x<=5 for x in b):raise ValueError('beta_norms 0..5')
 if len(set(b))!=len(b):raise ValueError('distinct beta samples required')
 if len(b)*(c['kmesh_n']**2+8)+3*c['path_points_per_segment']+1>500:raise ValueError('too many dense eigensolves')
 return c

def reciprocal_indices(cutoff):return np.array([(i,j) for i in range(-cutoff,cutoff+1) for j in range(-cutoff,cutoff+1)],int)
def eta_coefficients(c,differences):
 A=primitive(c['lattice'],1.);B=2*np.pi*np.linalg.inv(A).T;delta=np.asarray(differences);G=delta@B.T;r=c['hole_diameter_over_pitch']/2
 z=np.linalg.norm(G,axis=-1)*r;factor=np.ones_like(z);np.divide(2*j1(z),z,out=factor,where=z!=0)
 phase=sum(np.exp(-2j*np.pi*np.einsum('...i,i->...',delta,s)) for s in MOTIFS[c['lattice']])
 return (np.all(delta==0,axis=-1)/c['n_host']**2+(1/c['n_air']**2-1/c['n_host']**2)*np.pi*r*r/abs(np.linalg.det(A))*factor*phase)

def assemble(c,k_frac,beta_norm):
 validate(c);k=np.asarray(k_frac,float)
 if k.shape!=(2,) or not np.isfinite(k).all() or not np.isfinite(beta_norm) or not 0<=beta_norm<=5:raise ValueError('finite Bloch k, beta>=0 required')
 g=reciprocal_indices(c['pwe_cutoff']);B=2*np.pi*np.linalg.inv(primitive(c['lattice'],1.)).T
 q=np.column_stack(((g+k)@B.T,np.full(len(g),2*np.pi*beta_norm)))
 U=np.zeros((len(g),3,2))
 for i,v in enumerate(q):
  norm=np.linalg.norm(v)
  if norm==0:U[i,:,0]=[1,0,0];U[i,:,1]=[0,1,0]
  else:
   axis=np.eye(3)[np.argmin(abs(v))];u=np.cross(v,axis);u/=np.linalg.norm(u);U[i,:,0]=u;U[i,:,1]=np.cross(v/norm,u)
 curls=np.cross(q[:,:,None],U,axisa=1,axisb=1,axisc=1)
 eta=eta_coefficients(c,g[:,None,:]-g[None,:,:])
 M=(np.einsum('ipa,jpb->iajb',curls,curls)*eta[:,None,:,None]).reshape(2*len(g),2*len(g))
 err=float(np.linalg.norm(M-M.conj().T)/max(np.linalg.norm(M),1e-30))
 if err>1e-12:raise RuntimeError('non-Hermitian Maxwell operator')
 return M,q,U,g,err

def solve_point(c,k_frac,beta_norm,return_fields=False):
 M,q,U,g,herr=assemble(c,k_frac,beta_norm)
 vals,vec=eigh(M,subset_by_index=[0,c['bands']-1],driver='evr',check_finite=True)
 scale=max(np.linalg.norm(M),1.)
 if np.min(vals)<-1e-10*scale:raise RuntimeError('negative optical eigenvalue beyond roundoff')
 residual=float(np.max(np.linalg.norm(M@vec-vec*vals,axis=0)/(scale+abs(vals))))
 coeff=np.einsum('npa,nab->npb',U,vec.reshape(len(g),2,-1))
 divergence=float(np.max(abs(np.einsum('np,npb->nb',q,coeff)))/max(np.linalg.norm(q)*np.linalg.norm(coeff),1e-30))
 r={'frequencies_f':(np.sqrt(np.maximum(vals,0))/(2*np.pi)).tolist(),'residual_normwise':residual,'divergence_relative':divergence,'hermitian_relative':herr,'plane_waves':len(g),'transverse_dofs':len(M)}
 if return_fields:r.update(H_coefficients=coeff,G_indices=g,q=q)
 return r

def field_at(c,solution,k_frac,points_fractional):
 u=np.asarray(points_fractional,float);phase=np.exp(2j*np.pi*u@(solution['G_indices']+np.asarray(k_frac)).T)
 return np.einsum('pn,ncb->pcb',phase,solution['H_coefficients'])

def analyze(c):
 validate(c);cache={}
 def point(k,b):
  key=(tuple(k),float(b))
  if key not in cache:cache[key]=solve_point(c,k,b)
  return cache[key]['frequencies_f']
 ks,names=high_symmetry_path(c['lattice'],c['path_points_per_segment']);b=c['beta_norms'][0]
 path=[point(k,b) for k in ks];B=2*np.pi*np.linalg.inv(primitive(c['lattice'],c['pitch_um'])).T
 distance=[0.]
 for a,z in zip(ks,ks[1:]):distance.append(distance[-1]+float(np.linalg.norm(B@(np.asarray(z)-a))*c['pitch_um']/(2*np.pi)))
 kp=_kmesh_points(c['kmesh_n'],c['lattice']);gaps=[];summaries=[];edges=[]
 for beta in c['beta_norms']:
  spectrum=np.asarray([point(k,beta) for k in kp]);lower=spectrum.max(axis=0);upper=spectrum.min(axis=0)
  for i,(lo,hi) in enumerate(zip(lower[:-1],upper[1:])):
   edges.append({'beta_norm':beta,'band_below':i+1,'lower_edge':float(lo),'upper_edge':float(hi)})
   if hi>lo+1e-8 and lo>0:
    gaps.append({'band_below':i+1,'band_above':i+2,'f_low':float(lo),'f_high':float(hi),'width':float(hi-lo),'relative_width':float((hi-lo)/((hi+lo)/2)),
                 'beta_norm':beta,'beta_over_k0_at_f_low':float(beta/lo),'beta_over_k0_at_f_high':float(beta/hi),'status':'SAMPLED_GAP_CANDIDATE'})
  summaries.append({'beta_norm':beta,'sampled_k_count':len(kp),'gap_candidates':sum(x['beta_norm']==beta for x in gaps)})
 maskcfg=dict(c,schema='LSCNS_STEP5_BLOCH_INPUT_V083',beta_norms=[.001],bands=2,kmesh_n=2,path_points_per_segment=2)
 # Raster mask is only a web geometry preview; the Maxwell operator used analytic circles.
 mask=hole_mask(maskcfg)
 return {'schema':OUTPUT,'engine_version':'0.8.7','status':'VECTOR_PWE_SAMPLED_UNVERIFIED','input':copy.deepcopy(c),
         'input_sha256':hashlib.sha256(json.dumps(c,sort_keys=True,separators=(',',':')).encode()).hexdigest(),
         'model':'full-vector transverse H planewaves, curl eps^-1 curl; fixed beta, real non-dispersive epsilon',
         'unit_convention':'f=omega*pitch/(2*pi*c)=pitch/lambda, beta_norm=beta*pitch/(2*pi), k_frac reciprocal basis',
         'k_phase_convention':'H(r+a_i)=exp(i kB.a_i)H(r)','lattice_vectors_um':primitive(c['lattice'],c['pitch_um']).tolist(),'reciprocal_vectors_um_inv':B.tolist(),
         'basis_sites_fractional':[list(x) for x in MOTIFS[c['lattice']]],'unit_cell_air_mask':mask.astype(int).tolist(),
         'air_fill_fraction_analytic':float(len(MOTIFS[c['lattice']])*np.pi*(c['hole_diameter_over_pitch']/2)**2/abs(np.linalg.det(primitive(c['lattice'],1.)))),
         'path':{'k_frac':ks,'labels':names,'distance_norm':distance,'bands_f':path,'beta_norm':b},
         'gap_map':{'candidates':gaps,'band_edges':edges,'beta_summaries':summaries,'sampled_k_points':kp,'light_lines_beta_over_k0':[{'label':'air','n':c['n_air']},{'label':'host','n':c['n_host']}],'status':'SAMPLED_ONLY_NOT_COMPLETE_BZ_PROOF'},
         'checks':{'maximum_eigen_residual':max(v['residual_normwise'] for v in cache.values()),'maximum_divergence':max(v['divergence_relative'] for v in cache.values()),
                   'vector_maxwell':'IMPLEMENTED_TRANSVERSE_PWE','cutoff_convergence':'NOT_RUN','k_sampling_convergence':'NOT_RUN','complete_bandgap_certification':'NOT_VERIFIED','material_dispersion':'NOT_MODELED','MPB_external_comparison':'NOT_RUN'}}

def convergence(c,cutoffs=(2,3,4),kmeshes=(3,5,7)):
 rows=[]
 for cut in cutoffs:
  d=copy.deepcopy(c);d['pwe_cutoff']=cut;r=analyze(d);rows.append({'cutoff':cut,'kmesh_n':d['kmesh_n'],'band_edges':r['gap_map']['band_edges'],'candidates':r['gap_map']['candidates']})
 for count in kmeshes:
  d=copy.deepcopy(c);d['pwe_cutoff']=max(cutoffs);d['kmesh_n']=count;r=analyze(d);rows.append({'cutoff':max(cutoffs),'kmesh_n':count,'band_edges':r['gap_map']['band_edges'],'candidates':r['gap_map']['candidates']})
 return {'schema':'LSCNS_step5_convergence_v087','input':c,'cases':rows,'status':'CONVERGENCE_DATA_NOT_COMPLETE_GAP_CERTIFICATION'}

def defect_handoff(result,index,rings=4,defect_cells=7,core_surround_um=.15,domain_radius_um=None):
 from step5_bloch import defect_handoff as legacy,SCHEMA_OUTPUT
 if result.get('schema')!=OUTPUT:raise ValueError('computed vector PWE result required')
 r=copy.deepcopy(result);r['schema']=SCHEMA_OUTPUT
 d=legacy(r,index,rings,defect_cells,core_surround_um);gap=r['gap_map']['candidates'][index];f=(gap['f_low']+gap['f_high'])/2
 if defect_cells==19 and rings<3:raise ValueError('19-cell defect needs at least 3 rings for any remaining cladding holes')
 d['domain_radius_um']=float(domain_radius_um) if domain_radius_um is not None else (rings+3)*result['input']['pitch_um']
 d['metadata']['domain_radius_basis']='explicit numerical window or (rings+3)*pitch; requires independent window/PML convergence'
 from schema import validate_design
 validate_design(d)
 d['target_neff']=gap['beta_norm']/f
 if not 0<d['target_neff']<3.5:raise ValueError('gap beta/k0 outside finite solver target range')
 d['metadata'].update(provenance='actual vector PWE sampled gap -> triangular finite PBGF proposal',target_beta_over_k0=d['target_neff'],limitations='finite k/cutoff candidate; finite core mode, localization and independent PML/mesh loss validation still required')
 return d
