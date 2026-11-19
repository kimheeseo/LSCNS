import copy,json,sys,unittest,hashlib
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from lswave.reports import stamp,verify
from lswave.postprocess.formulas import loss,area,gamma,REGISTRY
from lswave.materials import index
from physical_metrics import confinement_loss,nonlinear_gamma
class Structure8(unittest.TestCase):
 def test_unicode_numeric_serialization_provenance_and_tamper(self):
  d={'name':'광섬유','x':1.,'small':1e-7,'nested':{'value':None}}
  r=stamp({'input':copy.deepcopy(d),'engine_version':'0.8.6'})
  self.assertTrue(verify(r));self.assertIn('광섬유',r['provenance']['canonical_input_json']);self.assertFalse(r['provenance']['physical_validation_upgraded'])
  r['input']['x']=2
  with self.assertRaises(ValueError):verify(r)
 def test_run_options_are_part_of_manifest_hash(self):
  r={'input':{'wavelength_nm':1550},'input_manifest':{'design':{'wavelength_nm':1550},'run_options':{'target_neff':1.44}}};stamp(r,r['input_manifest'],'0.8.6');self.assertTrue(verify(r));r['input_manifest']['run_options']['target_neff']=1.45
  with self.assertRaises(ValueError):verify(r)
 def test_missing_and_modified_hash_never_pass(self):
  with self.assertRaises(ValueError):verify({'input':{}})
  r=stamp({'input':{},'engine_version':'0.8.6'});r['input_sha256']='0'*64
  with self.assertRaises(ValueError):verify(r)
 def test_legacy_formula_wrappers_preserve_units_and_sign(self):
  expected=20/__import__('math').log(10)*(2*__import__('math').pi/1.55e-6)*(-1e-8)
  self.assertAlmostEqual(loss(-1e-8,1550),expected,places=12);self.assertEqual(confinement_loss(-1e-8,1550)['alpha_total_dB_m'],loss(-1e-8,1550));self.assertEqual(confinement_loss('-1e-8','1550')['alpha_total_dB_m'],loss(-1e-8,1550))
  self.assertEqual(nonlinear_gamma(2e-20,1550,80)['gamma_W_inv_m_inv'],gamma(2e-20,1550,80));self.assertEqual(area(4,2),8)
 def test_material_facade_keeps_explicit_legacy_air_assumption(self):
  self.assertEqual(index({'model':'air'},1550).real,1.00027)
  self.assertEqual(index({'model':'constant','n':1.},1550),1.)
 def test_actual_computed_example_contract_and_integration_scope(self):
  p=Path(__file__).resolve().parents[1]/'benchmarks/step8_v0810/example_mode_result.json';r=json.loads(p.read_text());self.assertTrue(verify(r));self.assertEqual(r['solver_version'],'0.8.6');self.assertIn('material_map',r);self.assertEqual(len(r['modes']),2)
  self.assertTrue(all(k in REGISTRY['formulas'] for k in ('loss','Aeff','gamma','D','B','ARROW','V_PCF')))
  for b in r['calculation_basis']:
   self.assertIn('source',b)
   if b['formula_id']=='Aeff':
    v=b['substituted_values'];self.assertGreater(v['actual_Gauss_points'],1000);self.assertEqual(b['result'],area(v['P'],v['Q']));self.assertFalse(v['PML_included'])
  self.assertFalse(r['precision_integration']['sparse_display_samples_integrated'])
if __name__=='__main__':unittest.main()
