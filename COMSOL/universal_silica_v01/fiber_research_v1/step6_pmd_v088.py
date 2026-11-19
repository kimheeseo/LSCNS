"""STEP6f uniform uncoupled birefringent fiber path (NOT statistical PMD).
Signed delta_n=nx-ny. delta_ng=delta_n-lambda*d(delta_n)/d(lambda), same lambda units.
DGD[ps]=length_m*abs(delta_ng)/299792458*1e12. B=abs(delta_n) alone is insufficient.
Random-coupling PMD[ps/sqrt(km)] additionally needs longitudinal correlation/coupling
statistics, twist, stress distributions and independently validated Jones propagation.
"""
import numpy as np
C=299792458.

def deterministic_dgd(wavelength_nm,signed_delta_n,length_m,tracking_evidence):
 lam=np.asarray(wavelength_nm,float);dn=np.asarray(signed_delta_n,float)
 if lam.ndim!=1 or dn.shape!=lam.shape or len(lam)<5 or not np.isfinite(lam).all() or not np.isfinite(dn).all() or np.any(np.diff(lam)<=0) or np.any(lam<=0) or not np.isfinite(length_m) or length_m<=0:raise ValueError('5+ increasing finite wavelengths, signed nx-ny and positive length needed')
 if not isinstance(tracking_evidence,dict) or tracking_evidence.get('polarization_branches_tracked') is not True or not tracking_evidence.get('source'):raise ValueError('explicit signed polarization branch tracking evidence required; sorted |B| cannot substitute')
 derivative=np.gradient(dn,lam,edge_order=2);ng=dn-lam*derivative
 return {'schema':'LSCNS_step6_dgd_v088','schema_version':'1.0.0','engine_version':'0.8.8','wavelength_nm':lam.tolist(),'signed_delta_n':dn.tolist(),
         'delta_ng':ng.tolist(),'DGD_ps':(length_m*abs(ng)/C*1e12).tolist(),'length_m':length_m,'tracking_evidence':tracking_evidence,
         'status':'UNIFORM_UNCOUPLED_FIBER_ESTIMATE','statistical_PMD_ps_sqrt_km':None,'certified':False,
         'limitations':['finite-difference derivatives need wavelength-step convergence','uniform axes and no longitudinal mode coupling assumed','existing Maxwell PMD statistics remain demonstration only; B does not determine PMD coefficient']}
