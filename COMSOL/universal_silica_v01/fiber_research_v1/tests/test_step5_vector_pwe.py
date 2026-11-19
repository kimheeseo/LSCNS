import copy,unittest
import numpy as np
from step5_vector_pwe import *
from step5_reference import reference_te_tm

class VectorPWE(unittest.TestCase):
 def test_uniform_folded_double_polarization_all_lattices(self):
  for lattice in MOTIFS:
   c=example_config();c.update(lattice=lattice,hole_diameter_over_pitch=.2,n_host=1.45,n_air=1.45,bands=8)
   for beta in (0.,.4,1.2):
    k=np.array([.13,.21]);r=solve_point(c,k,beta);g=reciprocal_indices(c['pwe_cutoff']);B=np.linalg.inv(primitive(lattice,1.)).T
    expected=np.sort(np.repeat(np.sqrt(np.sum(((g+k)@B.T)**2,axis=1)+beta**2)/1.45,2))[:8]
    self.assertLess(max(abs(np.array(r['frequencies_f'])-expected)),1e-10);self.assertLess(r['residual_normwise'],1e-10);self.assertLess(r['divergence_relative'],1e-12)
 def test_bloch_boundary_phase_and_hermiticity(self):
  c=example_config();c['hole_diameter_over_pitch']=.4;k=[.13,.21];r=solve_point(c,k,.5,True);p=np.array([[.15,.32],[.36,.22]])
  f=field_at(c,r,k,p)
  for axis in (0,1):
   shifted=p.copy();shifted[:,axis]+=1
   self.assertTrue(np.allclose(field_at(c,r,k,shifted),np.exp(2j*np.pi*k[axis])*f,atol=1e-10))
  self.assertLess(r['hermitian_relative'],1e-12)
 def test_nested_variational_cutoff(self):
  c=example_config();c.update(lattice='square',hole_diameter_over_pitch=.42,bands=6)
  values=[]
  for n in (2,3,4):c['pwe_cutoff']=n;values.append(np.array(solve_point(c,[.13,.21],.4)['frequencies_f']))
  self.assertTrue(np.all(values[1]<=values[0]+1e-12));self.assertTrue(np.all(values[2]<=values[1]+1e-12))
 def test_inverse_epsilon_fourier_against_independent_area_sum(self):
  c=example_config();c.update(lattice='square',hole_diameter_over_pitch=.4)
  a=(np.arange(512)+.5)/512;u,v=np.meshgrid(a,a,indexing='ij');mask=np.minimum(u,1-u)**2+np.minimum(v,1-v)**2<.2**2
  eta=np.where(mask,1.,1/1.45**2)
  for d in ([0,0],[1,0],[2,1]):
   numeric=np.mean(eta*np.exp(-2j*np.pi*(d[0]*u+d[1]*v)));self.assertLess(abs(numeric-eta_coefficients(c,np.array(d))),5e-4)
 def test_independent_real_space_te_tm_reference(self):
  c=example_config();c.update(lattice='square',hole_diameter_over_pitch=.42,bands=6,pwe_cutoff=6)
  p=np.array(solve_point(c,[.13,.21],0)['frequencies_f']);r=reference_te_tm(c,[.13,.21],64)
  self.assertLess(max(abs(p-r)/r),.02)
 def test_real_result_schema_gap_and_handoff_target(self):
  c=example_config();r=analyze(c);self.assertEqual(r['schema'],OUTPUT);self.assertFalse(r['checks']['complete_bandgap_certification']=='PASS')
  self.assertGreater(len(r['gap_map']['candidates']),0)
  d=defect_handoff(r,0);g=r['gap_map']['candidates'][0];self.assertAlmostEqual(d['target_neff'],g['beta_norm']/((g['f_low']+g['f_high'])/2))
  r['input']['lattice']='square'
  with self.assertRaises(NotImplementedError):defect_handoff(r,0)
 def test_nonoverlap_and_resource_validation(self):
  c=example_config();c['lattice']='kagome'
  with self.assertRaises(ValueError):validate(c)
  c=example_config();c['pwe_cutoff']=True
  with self.assertRaises(ValueError):validate(c)
if __name__=='__main__':unittest.main()
