import unittest,sys,copy
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from validation_csv_v089 import parse_comsol_csv,compare_tables,complex_cell,vector_to_table
from validation_reference_v089 import references
from validate_step7_v089 import unit_evidence,csv_evidence,item
class Validation7(unittest.TestCase):
 def test_actual_vector_JSON_conversion_keeps_mode_order_explicit(self):
  r=vector_to_table({'schema':'LSCNS_fiber_research_vector_v1','engine_version':'0.8.6','input':{'wavelength_nm':1550},'modes':[{'neff_real':1.44,'neff_imag':3e-8}]})
  self.assertEqual(r['rows'][0]['mode_id'],'1');self.assertEqual(r['source_solver_version'],'0.8.6');self.assertFalse(r['physical_cross_validation_certified'])
 def test_units_independent_identities(self):self.assertTrue(all(x['passed'] for x in unit_evidence()))
 def test_complex_exponents_pure_imaginary_and_parentheses(self):
  for text,value in [('1.2-3.4D-8i',(1.2,-3.4e-8)),('(1.2 + 3e-8i)',(1.2,3e-8)),('1.234e-8i',(0,1.234e-8)),('-1.234e-8j',(0,-1.234e-8)),('1.44',(1.44,0))]:self.assertEqual(complex_cell(text),value)
  for text in ['(1.4+1e-8i','1e309','float("nan")','1+2i+3i']:
   with self.assertRaises(ValueError):complex_cell(text)
 def test_comsol_comments_D_exponents_um_and_sign(self):
  t='% Model: numerical fixture\n% lambda (um);mode_id;real(ewfd.neff) (1);imag(ewfd.neff) (1)\n1.55;HE11;1.445;-3.0D-8\n'
  r=parse_comsol_csv(t,-1);self.assertEqual(r['rows'][0]['wavelength_nm'],1550);self.assertEqual(r['rows'][0]['neff_imag'],3e-8);self.assertEqual(r['rows'][0]['raw_neff_imag'],-3e-8)
 def test_missing_sign_unit_identity_and_imag_cannot_be_guessed(self):
  cases=[('lambda,mode,neff\n1550,a,1.44\n',{}),('wavelength_nm,neff\n1550,1.44\n',{}),('mode,neff_real\na,1.44\n',{'default_wavelength_nm':1550})]
  for s,kw in cases:
   with self.assertRaises(ValueError):parse_comsol_csv(s,1,**kw)
  with self.assertRaises(ValueError):parse_comsol_csv('mode,neff\na,1.44\n',0,1550)
 def test_single_column_row_mapping_requires_confirmation(self):
  t='% ewfd.neff (1)\n1.445-3e-8i\n1.44-4e-8i\n'
  with self.assertRaises(ValueError):parse_comsol_csv(t,-1,1550)
  r=parse_comsol_csv(t,-1,1550,confirm_row_mode_ids=True);self.assertEqual([x['mode_id'] for x in r['rows']],['1','2'])
 def test_duplicate_and_invalid_values_rejected(self):
  for t in ['mode,neff\na,1.44\na,1.44\n','mode,neff_real,neff_imag\na,NaN,1e-8\n','mode,neff\na,1.4+__import__("os")i\n']:
   with self.assertRaises(ValueError):parse_comsol_csv(t,1,1550)
 def test_zero_reference_imag_uses_absolute_tolerance(self):
  a=parse_comsol_csv('wavelength_nm,mode_id,neff\n1550,a,1.44+1e-10i\n',1);b=parse_comsol_csv('wavelength_nm,mode_id,neff\n1550,a,1.44\n',1)
  self.assertEqual(compare_tables(a,b,1,.05,0)['status'],'FAIL');r=compare_tables(a,b,1,.05,1e-10);self.assertEqual(r['status'],'PASS');self.assertIsNone(r['rows'][0]['Im_error_relative']);self.assertFalse(r['physical_cross_validation_certified'])
 def test_missing_and_ambiguous_rows_never_pass(self):
  a=parse_comsol_csv('mode,neff\na,1.44\nb,1.44\n',1,1550);b=parse_comsol_csv('mode,neff\na,1.44\nc,1.44\n',1,1550)
  self.assertEqual(compare_tables(a,b,1,.05,0)['status'],'NEEDS_EVIDENCE')
 def test_analytic_reference_not_substituted_for_missing_geometry(self):
  r=references();self.assertEqual(len(r['V2']['modes']),10);self.assertIsNone(r['V2']['exact_geometry_parameters']);self.assertIsNone(r['V3']['bent_complex_neff'])
  self.assertFalse(item('V2','x','NEEDS_EVIDENCE','x',{},r['V2'])['complete']);self.assertEqual(csv_evidence()['status'],'PASS')
if __name__=='__main__':unittest.main()
