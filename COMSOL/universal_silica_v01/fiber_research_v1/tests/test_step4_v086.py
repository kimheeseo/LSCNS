import copy,math,unittest
import numpy as np
from step4_v086 import integrate_mode,postprocess,dispersion,track_sweep,arrow_dispersive,detect_surface_pair
from physical_metrics import confinement_loss

def gaussian():
 x=np.linspace(-16,16,161);xx,yy=np.meshgrid(x,x);p=np.column_stack((xx.ravel(),yy.ravel()));a=np.exp(-(xx**2+yy**2)/16).ravel();z=np.zeros_like(a)
 pack=lambda v:np.column_stack((v,z)).tolist()
 q={'points_um':p.tolist(),'weights_um2':np.full(len(a),.04).tolist(),'material_labels':np.where(p[:,0]>0,'glass','air').tolist()}
 q.update({k:pack(a if k in ('Ex','Hy') else z) for k in ('Ex','Ey','Ez','Hx','Hy','Hz')})
 return {'neff_real':1.4,'neff_imag':1e-8,'field_quadrature':q},{'wavelength_nm':1550.,'materials':{'glass':{'model':'constant','n':1.5},'air':{'model':'air'}}}

def raw(rotated=False,swap=False):
 runs=[]
 for i,lam in enumerate(np.arange(1510.,1600.,10.)):
  modes=[];t=i*.7 if rotated else 0
  for j in range(2):
   vec=[math.cos(t),math.sin(t),0] if j==0 else [-math.sin(t),math.cos(t),0];e=np.tile(vec,(3,1))
   modes.append({'neff_real':1.45+2e-12*(lam-1550)**3+(0 if rotated else j*.005),'neff_imag':0.,'tracking_fields':{'points_um':[[0,0],[1,0],[0,1]],'weights_um2':[1,1,1],'E':np.stack((e,e*0),axis=-1).tolist()}})
  if swap and i%2:modes.reverse()
  runs.append({'input':{'wavelength_nm':float(lam),'parameters':{},'materials':{}},'modes':modes})
 return {'runs':runs}
class Step4V086(unittest.TestCase):
 def test_gaussian_area_moments_and_si_power(self):
  m,d=gaussian();r=integrate_mode(m,d,{'material_roles':{'glass':'glass','air':'air'}})
  self.assertAlmostEqual(r['Aeff_Et_um2'],math.pi*16,places=7);self.assertAlmostEqual(r['MFD_second_moment_Et_um'],8.,places=7)
  self.assertAlmostEqual(sum(r['material_power_fractions'].values()),1.);self.assertAlmostEqual(r['power_W'],.5*math.pi*8*1e-12,places=20)
  self.assertIsNone(r['n2_eff_m2_W']);self.assertIsNone(r['MFD_PetermannII_um'])
 def test_quartic_n2_and_radial_pii(self):
  m,d=gaussian();rr=np.linspace(0,20,2001)
  r=integrate_mode(m,d,{'nonlinear_materials':{'glass':{'n2_m2_W':2e-20,'source':'SYNTHETIC TEST'},'air':{'n2_m2_W':0.,'source':'SYNTHETIC ZERO'}},'radial_profile':{'basis':'axisymmetric_scalar','source':'analytic Gaussian TEST','r_um':rr.tolist(),'amplitude':np.column_stack((np.exp(-rr*rr/16),rr*0)).tolist()}})
  self.assertAlmostEqual(r['n2_eff_m2_W'],r['quartic_material_weights']['glass']*2e-20);self.assertAlmostEqual(r['MFD_PetermannII_um'],8.,delta=.001)
  self.assertAlmostEqual(r['gamma_W_inv_m_inv'],2*np.pi*r['n2_eff_m2_W']/(1550e-9*math.pi*16e-12))
 def test_loss_units_signed_absorption(self):
  a=confinement_loss(1e-8,1550);v=20/math.log(10)*2*math.pi/1.55e-6*1e-8
  self.assertAlmostEqual(a['alpha_total_dB_m'],v);self.assertAlmostEqual(a['alpha_total_dB_km'],1000*v)
  m,d=gaussian();d['materials']['glass']['kappa']=1e-5;self.assertIsNone(integrate_mode(m,d)['alpha_confinement_candidate_dB_m'])
 def test_display_samples_pml_rejected(self):
  m,d=gaussian();m['field_quadrature']['material_labels'][0]='pml'
  with self.assertRaises(ValueError):integrate_mode(m,d)
  with self.assertRaises(ValueError):integrate_mode({'field_samples':{}},d)
 def test_permutations_and_rotated_degenerate_subspaces(self):
  a=track_sweep(raw(swap=True));self.assertEqual(len(a['curves']),2);self.assertTrue(all(min(c['overlap_power'])>.999 for c in a['curves']))
  b=track_sweep(raw(rotated=True,swap=True));self.assertEqual(len(b['curves']),1);self.assertEqual(b['curves'][0]['kind'],'degenerate_subspace_centroid');self.assertGreater(min(b['curves'][0]['overlap_power']),.999)
 def test_topology_and_mesh_change_blocked(self):
  d=raw();d['runs'][4]['modes'].pop()
  with self.assertRaises(ValueError):track_sweep(d)
  d=raw();d['runs'][3]['modes'][0]['tracking_fields']['points_um'][0]=[.001,0]
  with self.assertRaises(ValueError):track_sweep(d)
 def test_exact_dispersion_slope_zero_and_flat(self):
  x=np.arange(1500.,1601.,10.);n=1.45+2e-12*(x-1550)**3;r=dispersion(x,n,np.ones(len(x)-1))
  D=-x/299792458*1e15*12e-12*(x-1550);S=-1e15/299792458*12e-12*(2*x-1550)
  self.assertLess(max(abs(np.array(r['D_ps_nm_km'])-D)),1e-5);self.assertLess(max(abs(np.array(r['S_ps_nm2_km'])-S)),1e-6);self.assertAlmostEqual(r['zero_dispersion_lambda_nm'][0],1550.,places=4)
  self.assertEqual(dispersion(x,np.full(len(x),1.45),np.ones(len(x)-1))['zero_dispersion_lambda_nm'],[])
 def test_arrow_implicit_missing_materials(self):
  d={'parameters':{'wall_thickness_um':.5},'materials':{'silica':{'model':'constant','n':1.5},'air':{'model':'constant','n':1.0}}};r=arrow_dispersive(d)
  self.assertAlmostEqual(r['resonance_lines'][0]['lambda_nm'],1000*math.sqrt(1.25),places=6)
  d['materials']['silica']={'model':'pure_silica'};self.assertLess(abs(arrow_dispersive(d)['resonance_lines'][0]['equation_residual_nm']),1e-6)
 def test_explicit_polarization_and_homer_context(self):
  m,d=gaussian();data={'schema':'LSCNS_fiber_research_vector_v1','input':d,'modes':[m,copy.deepcopy(m)]};data['modes'][1]['neff_real']+=1e-4;data['modes'][1]['neff_imag']=1e-7
  r=postprocess(data,{'polarization_pair':[0,1],'polarization_identity_source':'SYNTHETIC x/y','mode_identities':{'0':'FM','1':'HOM'},'identities_source':'SYNTHETIC LABELS','losses_converged':True});self.assertAlmostEqual(r['birefringence']['B'],1e-4);self.assertAlmostEqual(r['HOMER']['HOMER_ratio'],10.)
 def test_surface_avoided_crossing_requires_field_exchange(self):
  x=np.linspace(1500.,1600.,9);det=(x-1550)*1e-5;gap=np.sqrt(det*det+1e-8)
  a={'kind':'individual_mode','neff_real':(1+gap/2).tolist()};b={'kind':'individual_mode','neff_real':(1-gap/2).tolist()}
  tracked={'parameter':'wavelength_nm','parameter_values':x.tolist(),'curves':[a,b]}
  fg=(.5+.5*det/gap).tolist();fa=(1-np.array(fg)).tolist()
  r=detect_surface_pair(tracked,0,1,[fg,fa],'synthetic two-state coupled-mode model TEST')
  self.assertEqual(len(r['candidates']),1);self.assertAlmostEqual(r['candidates'][0]['lambda_nm'],1550.)
  self.assertEqual(detect_surface_pair(tracked,0,1,[fg,fg],'synthetic TEST')['candidates'],[])
 def test_parameter_sweep_no_fake_dispersion(self):
  d=raw()
  for i,r in enumerate(d['runs']):r['input']['wavelength_nm']=1550.;r['input']['parameters']['core_radius_um']=3.+i*.1
  d['parameter']='parameters.core_radius_um';d['parameter_values']=[3.+i*.1 for i in range(len(d['runs']))]
  result=track_sweep(d);self.assertNotIn('D_ps_nm_km',result['curves'][0])
 def test_missing_nonlinear_source_is_not_silently_filled(self):
  m,d=gaussian();r=integrate_mode(m,d,{'nonlinear_materials':{'glass':{'n2_m2_W':2e-20},'air':{'n2_m2_W':0.}}})
  self.assertIsNone(r['gamma_W_inv_m_inv'])
if __name__=='__main__':unittest.main()
