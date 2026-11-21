import unittest
from run_physics_v6 import convergence,relative_mode_candidates
import test_physics_modal as modal_test


class ConvergenceEvidence(unittest.TestCase):
    def test_field_family_projection_does_not_certify_hom(self):
        helper=modal_test.ModalEvidence();a,b=helper.mode(0,0),helper.mode(1,1)
        c=helper.mode(0,2,True);c['neff_real']=.98
        result=relative_mode_candidates([a,c],[a,b])
        self.assertEqual(result[0]['label'],'FM_SUBSPACE_MATCH')
        self.assertEqual(result[1]['label'],'HOM_CANDIDATE_NOT_IDENTIFIED')

    def test_small_residual_is_not_convergence(self):
        self.assertFalse(convergence([1.377984,1.379063,1.380274],1e-5))

    def test_two_final_changes_and_finite_data_required(self):
        self.assertFalse(convergence([1.,1.+1e-7],1e-5))
        self.assertFalse(convergence([1.,1.01,1.01000001],1e-5))
        self.assertFalse(convergence([1.,float('nan'),1.],1e-5))
        self.assertTrue(convergence([1.,1.+1e-7,1.+2e-7],1e-5))


if __name__=='__main__':unittest.main()
