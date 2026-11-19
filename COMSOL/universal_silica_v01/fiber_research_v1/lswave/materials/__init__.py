"""Material index/provenance facade; no new assumed physical constants."""
def index(material,wavelength_nm):
 from materials import index as existing
 return existing(material,wavelength_nm)
def provenance(material):
 from materials import provenance as existing
 return existing(material)
