"""Browser contract facade. Static ES modules remain at ../../../web/."""
def result_contract():
 return {'version':'1.1.0','schema':'LSCNS_fiber_research_vector_v1','required_provenance':['input_sha256','solver_version','timestamp_utc','provenance.canonical_input_json'],'display_fields':['field_samples','material_map','calculation_basis'],'browser_computation_scope':'scalar/axisymmetric legacy only; vector FEM and Bloch eigenproblems run in Python','legacy_result_policy':'display with missing-provenance warning; never silently certify'}
