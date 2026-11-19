"""Stable vector-mode API: original assembly/spectrum, canonical basis/provenance.
Full Gauss data is used in Python and omitted by default after aggregation.
Sparse field_samples remain visualization only. material_map is labeled base
material n, not an anisotropic tensor or an inferred effective cladding index.
"""
import copy,json
import numpy as np
from ..reports import stamp
from ..postprocess.formulas import loss,area,basis

def modes(design,mesh=None,retain_quadrature=False,**options):
 from geometry import build_mesh
 from vector_fem import solve
 from materials import index,provenance
 from step4_v086 import quadrature
 if mesh is None:mesh=build_mesh(design)
 options=dict(options);options['export_full_quadrature']=True
 result=solve(design,mesh,**options);entries=[]
 for i,m in enumerate(result['modes']):
  alpha=loss(m['neff_imag'],design['wavelength_nm']);entries.append(dict(mode_index=i,**basis('loss',{'lambda_nm':design['wavelength_nm'],'Im(neff)':m['neff_imag']},alpha)))
  p,w,f,labels=quadrature(m);I=abs(f['Ex'])**2+abs(f['Ey'])**2;P=float(I@w);Q=float((I*I)@w)
  entries.append(dict(mode_index=i,**basis('Aeff',{'P':P,'Q':Q,'PML_included':False,'actual_Gauss_points':len(w)},area(P,Q))))
  if not retain_quadrature:m.pop('field_quadrature',None)
 centers=np.asarray(mesh['points_um'])[np.asarray(mesh['triangles'])].mean(axis=1);labels=np.asarray(mesh['material_ids']);ids=np.flatnonzero(labels!='pml')
 if len(ids)>1800:ids=ids[np.linspace(0,len(ids)-1,1800,dtype=int)]
 vals=[index(design['materials'][labels[j]],design['wavelength_nm']) for j in ids]
 result['material_map']={'points_um':centers[ids].tolist(),'n_real':[v.real for v in vals],'n_imag':[v.imag for v in vals],'material_labels':labels[ids].tolist(),'basis':'actual triangle centroids; base scalar material n from labels; physical domain, PML excluded; not stress/bend tensor or a smoothed n(r) inference','sources':{str(s):provenance(design['materials'][s]) for s in set(labels[ids])}}
 result['calculation_basis']=entries;result['precision_integration']={'performed_in':'Python actual physical-domain Gauss quadrature','full_quadrature_retained':retain_quadrature,'sparse_display_samples_integrated':False}
 def encode(v):
  if isinstance(v,(np.ndarray,np.generic)):return v.tolist()
  raise TypeError('run option is not JSON serializable')
 manifest_options=json.loads(json.dumps(dict(options,retain_quadrature=retain_quadrature),default=encode,allow_nan=False))
 result['input_manifest']={'design':copy.deepcopy(design),'run_options':manifest_options}
 return stamp(result,result['input_manifest'],'0.8.6')
