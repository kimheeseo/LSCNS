"""Verified reference facts; absent parameters remain None. No inferred geometry.
COMSOL Table1 passive Im sign converted from exp(+i wt) to exp(-i wt) explicitly.
V3 axisymmetric neff*r0 eigenvalue must NOT be compared as dimensionless neff.
"""
V2_URL='https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.microstructured_optical_fiber/microstructured_optical_fiber.html'
V3_URL='https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.step_index_fiber_bend/step_index_fiber_bend.html'
FSM_URL='https://arxiv.org/pdf/physics/0307010'
ARROW_URL='https://doi.org/10.3390/fib6040089'

def references():
 names=['HE11_a','HE11_b','TE01','HE21_a','HE21_b','TM01','HE31_a','EH11_a','EH11_b','HE31_b']
 real=[1.44539464,1.44539457,1.43858017,1.43844305,1.43844288,1.43836454,1.43039731,1.42995211,1.42995194,1.42925271]
 im=[3.18614991e-8,3.18633520e-8,5.33455430e-7,9.69173128e-7,9.69308091e-7,1.40665012e-6,2.15696869e-5,1.58836822e-5,1.58836249e-5,8.69622507e-6]
 return {'verified_date':'2026-11-20','V2':{'title':'Leaky Modes in a Microstructured Optical Fiber — COMSOL 6.4 Table1 PML','url':V2_URL,
 'modes':[{'mode_id':n,'neff_real':r,'neff_imag':v} for n,r,v in zip(names,real,im)],'wavelength_nm':None,'exact_geometry_parameters':None,
 'missing':'original microstructured_optical_fiber_parameters.txt: dimensions, indices, wavelength, physical/PML sizes; no values inferred from figures',
 'conversion':'published negative Im multiplied by -1, passive convention exp(i beta z-i omega t)',
 'original_method_reference':'Uranus & Hoekstra, Optics Express 12 (2004), DOI 10.1364/OPEX.12.002795'},
 'V3':{'title':'Step-Index Fiber Bend — COMSOL 6.4','url':V3_URL,'wavelength_nm':1550,'n_core':1.4457,'n_clad':1.4378,'core_radius_um':8.,'straight_cladding_radius_um':40.,'bend_radius_mm':3.,'bent_square_side_um':100.,'pml_thickness_um':20.,
 'straight_neff_rounded':1.4444,'bent_complex_neff':None,'status':'NEEDS_EVIDENCE','missing':'full precision axisymmetric bent complex neff and r0 normalization, equivalent formulation/window comparison; average mode radius is about R+4 um, not exact input'},
 'V6':{'title':'Modal cut-off and the V parameter in photonic crystal fibers','url':FSM_URL,'n_glass':1.444,'n_air':1.,'V_cutoff':3.141592653589793,'threshold_d_over_pitch':.406,'formula':'V_PCF=(2*pi*Lambda/lambda)*sqrt(n_FM^2-n_FSM^2)','scope':'triangular one-hole defect; finite wavelength samples alone cannot prove endlessly single mode'},
 'V7':{'title':'A Method to Process Hollow-Core Anti-Resonant Fibers into Fiber Filters (2018)','url':ARROW_URL,'formula':'lambda_m=2*t*sqrt(ng^2-na^2)/m','scope':'ARROW approximate resonance; FEM spectral peak needs mode tracking and independent mesh/PML convergence'}}
