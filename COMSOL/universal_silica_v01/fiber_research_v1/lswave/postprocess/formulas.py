"""Canonical research scalar formulas. Units and provenance live in formulas.json.
Legacy wrappers delegate here; no source string is evaluated as executable code.
"""
import json,math
from pathlib import Path
import numpy as np
REGISTRY=json.loads(Path(__file__).with_name('formulas.json').read_text())
def _finite(value,name,positive=False,nonnegative=False):
 if isinstance(value,bool) or not np.isscalar(value) or not np.isfinite(value) or positive and value<=0 or nonnegative and value<0:raise ValueError('invalid '+name)
 return float(value)
def loss(imag,wavelength_nm):
 imag=_finite(imag,'Im(neff)');wavelength_nm=_finite(wavelength_nm,'wavelength_nm',positive=True)
 return 20/math.log(10)*(2*math.pi/(wavelength_nm*1e-9))*imag
def area(P,Q):
 if not np.isfinite([P,Q]).all() or min(P,Q)<=0:raise ValueError('finite positive physical intensity integrals required')
 return P*P/Q
def gamma(n2,wavelength_nm,Aeff_um2):
 n2=_finite(n2,'n2',nonnegative=True);wavelength_nm=_finite(wavelength_nm,'wavelength_nm',positive=True);Aeff_um2=_finite(Aeff_um2,'Aeff_um2',positive=True)
 return 2*math.pi*n2/(wavelength_nm*1e-9*Aeff_um2*1e-12)
def basis(key,inputs,result):
 spec=REGISTRY['formulas'][key]
 return {'formula_id':key,**spec,'substituted_values':inputs,'result':float(result),'registry_version':REGISTRY['schema_version']}
