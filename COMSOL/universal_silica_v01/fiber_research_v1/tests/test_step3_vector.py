"""STEP 3 independent references, gauge/degeneracy and open-boundary diagnostics."""
import copy
import json
import math
from pathlib import Path
import sys
import unittest
import numpy as np
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from modal_diagnostics import independent_indices,principal_correlations,weighted_columns,disk_probes
from modal_study import match_subspace,sensitivity_verdict,input_digest
from step3_reference import he11_exact,leaky_circular_exact
from validate_step3 import benchmark_design
from schema import validate_design


class AlgebraAndContract(unittest.TestCase):
    def test_degenerate_independent_modes_are_kept_and_phase_duplicate_removed(self):
        a=np.array([1,0,0],complex);b=np.array([0,1,0],complex)
        self.assertEqual(independent_indices(np.column_stack((a,1j*a,b))),[0,2])

    def test_rotated_complex_degenerate_basis_tracks_as_subspace(self):
        a=np.array([[1,0],[0,1],[0,0]],complex)
        unitary=np.array([[1,1j],[1j,1]])/math.sqrt(2)
        self.assertTrue(np.allclose(principal_correlations(a,a@unitary),[1,1]))
        self.assertEqual(principal_correlations(a,a[:,:1]),[])

    def test_disk_probe_area_and_core_radial_break(self):
        p,w=disk_probes(6,1.3)
        self.assertAlmostEqual(w.sum(),math.pi*36,places=10)
        self.assertTrue(np.all(np.linalg.norm(p,axis=1)<6))

    def test_no_loss_certification_from_zero_or_sign_unstable_imaginary_part(self):
        rows=[dict(state='COMPUTED',tracking={'status':'PASS_TRACKING'},neff_real=[1.4,1.4],neff_imag=[x,x]) for x in (-1e-12,0,1e-12)]
        r=sensitivity_verdict(rows)
        self.assertEqual(r['status'],'UNRESOLVED_NEAR_ZERO_LOSS');self.assertFalse(r['loss_certified'])

    def test_opposite_imaginary_parts_cannot_cancel_in_a_pair_average(self):
        rows=[dict(state='COMPUTED',tracking={'status':'PASS_TRACKING'},neff_real=[1.4,1.4],neff_imag=[-.004,.004]) for _ in range(3)]
        result=sensitivity_verdict(rows)
        self.assertEqual(result['status'],'FAIL_SENSITIVITY')
        self.assertLess(result['minimum_Im'],0)
        self.assertFalse(result['loss_certified'])

    def test_trace_digest_changes_with_physical_and_numerical_inputs(self):
        d=benchmark_design();e=copy.deepcopy(d);e['solver']['element_order']=1
        self.assertNotEqual(input_digest(d),input_digest(e))
        e=copy.deepcopy(d);e['materials']['core']['n']=1.51
        self.assertNotEqual(input_digest(d),input_digest(e))

    def test_exact_references_have_small_characteristic_residuals(self):
        g=he11_exact(1.5,1.44,1.3,1.55)
        self.assertAlmostEqual(g['neff_real'],1.468523976355074,places=10)
        self.assertLess(abs(g['characteristic_residual']),1e-10)
        l=leaky_circular_exact(1,1.45,3,1.55,.981+.005j)
        self.assertGreater(l['neff_imag'],0);self.assertLess(l['characteristic_residual'],1e-10)

    def test_solver_options_reject_unsupported_and_nonfinite(self):
        for key,value in [('element_order',True),('backend','scalar'),('tolerance',float('nan')),('max_dofs',2),('sbc_iterations',7)]:
            d=benchmark_design();d['solver'][key]=value
            with self.subTest(key=key),self.assertRaises(ValueError):validate_design(d)


try:
    import gmsh,skfem
    DEPS=True
except (ImportError,OSError):DEPS=False


@unittest.skipUnless(DEPS,'requires actual gmsh/scikit-fem')
class ActualVectorChecks(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from geometry import build_mesh
        from vector_fem import solve
        cls.d=benchmark_design();cls.d['mesh'].update(curvature_segments=144)
        cls.mesh=build_mesh(cls.d)
        cls.p,cls.w=disk_probes(5.5,1.3,radial_order=8,angular_points=16)
        cls.result=solve(cls.d,cls.mesh,n_modes=2,target_neff=1.46852,max_sample_points=16,probe_points=cls.p,probe_weights=cls.w)

    def test_two_guided_modes_neff_gauss_and_6_components(self):
        r=self.result;self.assertEqual(r['engine_version'],'0.8.5');self.assertEqual(len(r['modes']),2)
        reference=he11_exact(1.5,1.44,1.3,1.55)['neff_real']
        for m in r['modes']:
            self.assertLess(abs(m['neff_real']-reference),1e-5)
            self.assertLess(m['weak_gauss_relative'],1e-7)
            self.assertLess(m['residual_relative'],1e-7)
            self.assertEqual(set(m['field_samples']['components']),set(('Ex','Ey','Ez','Hx','Hy','Hz')))
            self.assertAlmostEqual(m['classification']['power_W'],1.,places=10)
            self.assertTrue(m['classification']['finite_six_components'])
        self.assertFalse(r['validation']['loss_dB_km_certified'])
        json.dumps(r,allow_nan=False)

    def test_companion_and_beta_squared_are_equivalent_on_identical_mesh(self):
        from vector_fem import solve
        d=copy.deepcopy(self.d);d['solver']['backend']='companion'
        r=solve(d,self.mesh,n_modes=2,target_neff=1.46852,max_sample_points=8,probe_points=self.p,probe_weights=self.w)
        self.assertEqual(match_subspace(self.result,r)['status'],'PASS_TRACKING')
        for a,b in zip(sorted(m['neff_real'] for m in self.result['modes']),sorted(m['neff_real'] for m in r['modes'])):
            self.assertLess(abs(a-b),1e-9)
        self.assertEqual(r['solver_diagnostics']['pencil_dofs'],2*r['solver_diagnostics']['mixed_dofs'])

    def test_budget_and_nonfinite_mesh_are_rejected(self):
        from vector_fem import solve
        d=copy.deepcopy(self.d);d['solver']['max_dofs']=100
        with self.assertRaises(MemoryError):solve(d,self.mesh)
        bad=dict(self.mesh,points_um=self.mesh['points_um'].copy());bad['points_um'][0,0]=float('nan')
        with self.assertRaisesRegex(ValueError,'finite'):solve(self.d,bad)

    def test_pml_h_field_uses_transformed_inverse_mu(self):
        from vector_fem import _tensor
        d=benchmark_design('radiative')
        x=np.array([[[7.]],[[.5]]]);eps=np.ones((1,1),complex)
        _,_,_,_,gxx,gxy,gyy,gzz=_tensor(x,eps,d)
        r=float(np.linalg.norm(x));q=(r-6)/2
        sr=1+4j*q**3;st=1+4j*2*q**4/(4*r)
        self.assertAlmostEqual(complex(gzz[0,0]),1/(sr*st))
        self.assertFalse(np.allclose([gxx[0,0],gyy[0,0]],[1,1]))

    def test_nonzero_tiny_absorption_is_not_rounded_to_a_real_pencil(self):
        from vector_fem import solve
        d=copy.deepcopy(self.d)
        for material in d['materials'].values():material['kappa']=1e-20
        r=solve(d,self.mesh,n_modes=2,target_neff=1.46852,max_sample_points=8)
        self.assertFalse(r['solver_diagnostics']['real_pencil'])
        self.assertFalse(r['validation']['loss_dB_km_certified'])

    def test_sbc_uses_requested_target_and_can_iterate(self):
        from vector_fem import solve
        d=copy.deepcopy(self.d);d['boundary'].update(type='sbc',thickness_um=0)
        d['solver']['sbc_iterations']=3
        r=solve(d,self.mesh,n_modes=2,target_neff=1.46852,max_sample_points=8)
        history=r['solver_diagnostics']['sbc_iteration']['history']
        self.assertAlmostEqual(history[0]['boundary_target_real'],1.46852)
        self.assertGreaterEqual(len(history),2)
        self.assertFalse(r['validation']['loss_dB_km_certified'])


if __name__=='__main__':unittest.main()
