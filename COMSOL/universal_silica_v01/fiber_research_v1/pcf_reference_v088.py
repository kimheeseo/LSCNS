"""Verified literature inputs; reference fit is NOT numerical FEM validation V2/V6.
Folkenberg et al., physics/0307009 Table I and Eq.(1):
lambda_cut/Lambda = alpha*(d/Lambda-0.406)^gamma;
alpha=2.80 +/-0.12, gamma=0.89 +/-0.02; d/Lambda measurement +/-0.02.
Do not call a finite lossy PCF endlessly-single-mode from this reference alone.
"""
import math
SOURCE='https://arxiv.org/pdf/physics/0307009'

def cutoff_reference(pitch_um,diameter_ratio):
 if not math.isfinite(pitch_um) or pitch_um<=0 or not math.isfinite(diameter_ratio) or not .406<diameter_ratio<1:raise ValueError('reference fit requires pitch>0 and .406<d/Lambda<1')
 v=diameter_ratio-.406
 bounds=[a*v**g for a in (2.68,2.92) for g in (.87,.91)]
 return {'lambda_cut_um':pitch_um*2.8*v**.89,'fit_coefficient_interval_um':[pitch_um*min(bounds),pitch_um*max(bounds)],
 'source':SOURCE,'status':'PUBLISHED_FIT_NOT_FEM','note':'interval varies fit coefficients only; geometry +/-0.02 and experimental/systematic errors are separate; paper test ratios .54..68'}

def reference_audit():
 rows=[dict(fiber=i+1,pitch_um=p,diameter_ratio=d,**cutoff_reference(p,d)) for i,(p,d) in enumerate(zip([1.41,1.20,1.52,1.10,1.21,1.20,1.20],[.66,.64,.68,.63,.59,.57,.54]))]
 return {'schema':'LSCNS_pcf_reference_v088','V2':{'status':'BLOCKED_REFERENCE_GEOMETRY_PARAMETERS','source':'https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html',
 'reference_Re_neff_HE11':[1.44539464,1.44539457],'reference_abs_Im_neff_HE11':[3.18614991e-8,3.18633520e-8],
 'needed':'microstructured_optical_fiber_parameters.txt: exact hole/ring dimensions, indices, wavelength, physical domain and PML thickness; do not infer geometry from figures',
 'sign_note':'COMSOL table negative Im uses different convention; this solver passive Im is positive','passed':False},
 'V6':{'status':'REFERENCE_INPUTS_VERIFIED_FEM_CUTOFF_NOT_VALIDATED','passed':False,'source':SOURCE,'fibers':rows,
 'needed':'matched multi-ring geometry, wavelength-tracked FM/HOM with mesh/PML/loss convergence and cladding space-filling mode reference; fit alone cannot certify cutoff'}}
