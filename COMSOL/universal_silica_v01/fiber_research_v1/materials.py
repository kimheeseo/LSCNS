"""Verified pure silica Sellmeier and explicit unverified assumptions."""
import math
MALITSON = "https://doi.org/10.1364/JOSA.55.001205"
B=(.6961663,.4079426,.8974794)
C=(.0684043**2,.1162414**2,9.896161**2)
def silica_index(nm):
    w=nm*1e-3
    if not .21<w<3.7: raise ValueError("silica Sellmeier outside 0.21-3.7 um")
    return math.sqrt(1+sum(b*w*w/(w*w-c) for b,c in zip(B,C)))
def index(material,nm):
    model=material["model"]
    if model=="pure_silica": n=silica_index(nm)
    elif model=="air": n=1.00027
    elif model=="constant": n=float(material["n"])
    else: raise NotImplementedError("Ge/F doped Sellmeier coefficients TODO")
    return complex(n,float(material.get("kappa",0)))
def provenance(material):
    return MALITSON if material["model"]=="pure_silica" else ("Assumed constant air n=1.00027, ambient conditions unspecified" if material["model"]=="air" else material.get("source_url","User-defined constant index, not a dispersion fit"))

