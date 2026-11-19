"""Independent analytic identities and guards for STEP6."""
import sys,unittest,tempfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import numpy as np
from step6_splice_v088 import gaussian_grid,splice,fresnel
from step6_sweep_v088 import grid_sweep
from step6_pmd_v088 import deterministic_dgd,C
from step6_fields_v088 import polarization_pair,subspace_overlap
from pcf_reference_v088 import reference_audit,cutoff_reference
class Step6(unittest.TestCase):
 def test_gaussian_size_offset_polarization_fresnel(self):
  a=gaussian_grid(10);b=gaussian_grid(14)
  self.assertAlmostEqual(splice(a,a)['eta_Et'],1,places=12)
  self.assertAlmostEqual(splice(a,b)['eta_Et'],(2*5*7/(25+49))**2,places=11)
  self.assertAlmostEqual(splice(a,a,offset_um=(2,0))['eta_Et'],np.exp(-4/25),delta=.002)
  self.assertLess(splice(a,gaussian_grid(10,polarization=(0,1)))['eta_Et'],1e-25)
  self.assertAlmostEqual(fresnel(1,1.5),.96)
  self.assertAlmostEqual(splice(a,a,interface_indices=[1,1.5])['eta_with_requested_fresnel'],.96,places=11)
 def test_gaussian_angle_and_gap(self):
  a=gaussian_grid(10,side=193);tilt=.3;q=2*np.pi/1.55*np.sin(np.deg2rad(tilt))
  self.assertAlmostEqual(splice(a,a,angle_deg=(tilt,0))['eta_Et'],np.exp(-(q*5)**2/4),places=10)
  z=40;zr=np.pi*25/1.55
  self.assertAlmostEqual(splice(a,a,gap_um=z)['eta_Et'],1/(1+(z/(2*zr))**2),delta=.002)
 def test_shifted_aperture_not_renormalized(self):
  a=gaussian_grid(10);self.assertLess(splice(a,a,offset_um=(55,0))['eta_Et'],1e-15)
 def test_splice_rejects_invalid_inputs(self):
  a=gaussian_grid(10)
  for kw in ({'offset_um':(float('nan'),0)},{'gap_um':-1},{'angle_deg':(12,0)},{'gap_um':3,'gap_n':1,'interface_indices':[1.45,1.1,1]}):
   with self.assertRaises(ValueError):splice(a,a,**kw)
  with self.assertRaises(ValueError):splice(a,gaussian_grid(10,wavelength_nm=1310))
 def test_pareto_failures_constraints_resume_saved_design(self):
  d={'parameters':{'x':0,'y':0}};axes={'parameters.x':[0,1,2],'parameters.y':[0,1]};obj={'cost':{'direction':'min','unit':'USD'},'performance':{'direction':'max','unit':'arb'}}
  def evaluate(d):
   x=d['parameters']['x'];y=d['parameters']['y']
   if x==2 and y==1:raise RuntimeError('failed physical solve')
   return {'cost':x+y,'performance':x,'loss':y}
  with tempfile.TemporaryDirectory() as tmp:
   path=Path(tmp)/'partial.json';r=grid_sweep(d,axes,obj,evaluate,{'loss':{'max':0}},checkpoint=path,evaluator_id='analytic-v1')
   self.assertEqual(r['pareto_case_indices'],[0,2,4]);self.assertEqual(r['best_per_objective'],{'cost':0,'performance':4});self.assertEqual(r['cases'][5]['status'],'FAILED')
   rr=grid_sweep(d,axes,obj,lambda _:self.fail('must resume'),{'loss':{'max':0}},checkpoint=path,evaluator_id='analytic-v1');self.assertEqual(rr['best_designs'],r['best_designs'])
   with self.assertRaises(ValueError):grid_sweep(d,axes,obj,evaluate,checkpoint=path,evaluator_id='different')
 def test_signed_birefringence_dgd_identity(self):
  lam=np.linspace(1500,1600,9);dn=.0002+lam*2e-8;L=12
  r=deterministic_dgd(lam,dn,L,{'polarization_branches_tracked':True,'source':'analytic linear test'})
  np.testing.assert_allclose(r['delta_ng'],.0002,rtol=1e-12);np.testing.assert_allclose(r['DGD_ps'],L*.0002/C*1e12)
  self.assertIsNone(r['statistical_PMD_ps_sqrt_km'])
  with self.assertRaises(ValueError):deterministic_dgd(lam,dn,L,{})
 def test_rotated_degenerate_subspace_and_wrong_pm_pair(self):
  p=[[0,0],[1,0],[0,1],[1,1]];w=[1]*4
  def mode(field,pol,n):
   e=np.zeros((4,3),complex);e[:,:2]=field
   return {'tracking_fields':{'points_um':p,'weights_um2':w,'E':np.stack((e.real,e.imag),axis=-1).tolist()},'classification':{'polarization_class':pol},'neff_real':n,'status':'candidate_unverified'}
  x=np.array([[1,0]]*4);y=np.array([[0,1]]*4);a=mode(x,'quasi_x',1.45);b=mode(y,'quasi_y',1.4498)
  c=mode((x+y)/np.sqrt(2),'mixed',1.45);d=mode((x-y)/np.sqrt(2),'mixed',1.45)
  np.testing.assert_allclose(subspace_overlap([a,b],[c,d]),[1,1],atol=1e-14)
  r={'input':{'wavelength_nm':1550},'modes':[a,b]};self.assertAlmostEqual(polarization_pair(r)['B'],.0002)
  b['tracking_fields']['E'][:2]=[[[0,0],[0,0],[0,0]]]*2
  self.assertEqual(polarization_pair(r)['status'],'PAIR_REJECTED')
 def test_pcf_no_fake_fem_pass(self):
  r=reference_audit();self.assertFalse(r['V2']['passed']);self.assertFalse(r['V6']['passed']);self.assertEqual(len(r['V6']['fibers']),7)
  self.assertAlmostEqual(cutoff_reference(1,.606)['lambda_cut_um'],2.8*.2**.89)
if __name__=='__main__':unittest.main()
