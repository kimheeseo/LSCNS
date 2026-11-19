"""STEP6b Gmsh+mechanics+anisotropic HCurl Maxwell smoke. NOT V1/V4/V5 / COMSOL certification."""
import copy
import sys
import unittest
from pathlib import Path
import numpy as np
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from schema import load_presets
from vector_fem import _tensor
from pm_cli import derive_modal_birefringence

try:
    import gmsh
    import skfem
    HAVE_DEPS=True
except ImportError:
    HAVE_DEPS=False

SOURCE="https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.stress_optical/stress_optical.html"

def small_pm():
    d=copy.deepcopy(load_presets()["polarization_maintaining__panda"])
    d["domain_radius_um"]=12
    d["parameters"]["core_radius_um"]=2.2
    d["parameters"]["rod_offset_um"]=6.5
    d["parameters"]["rod_radius_um"]=2.0
    d["mesh"]["elements_per_wavelength"]=4
    d["mesh"]["bulk_max_um"]=1.0
    d["boundary"]["thickness_um"]=2.
    return d

def fake_props():
    # Pure numerical smoke material parameters; are NOT empirical silica/B2O3 constants.
    m={}
    for name in ("background","core","stress_rod"):
        m[name]={"E_Pa":8e10,"poisson_ratio":.2,
                 "alpha_1_per_K":(3e-6 if name=="stress_rod" else 5e-7),
                 "p11":.12,"p12":.27,
                 "mechanical_source_url":SOURCE,"photoelastic_source_url":SOURCE}
    return {"schema":"LSCNS_step6b_materials_v1","delta_T_K":-400.,"materials":m,
            "test_only":True}

class Step6bSmoke(unittest.TestCase):
    def test_analytic_anisotropic_epsilon_addition_and_straight_unaffected(self):
        d=small_pm()
        coords=np.array([[1.,2.],[0.,0.]])
        ep=np.ones((2,2),complex)*2.25
        base=_tensor(coords,ep,d)
        delta=(np.ones_like(ep)*.1,np.ones_like(ep)*.02,
               np.ones_like(ep)*(-.1),np.ones_like(ep)*.05)
        transformed=_tensor(coords,ep,d,delta)
        self.assertAlmostEqual(transformed[0][0,0].real,2.35)
        self.assertAlmostEqual(transformed[1][0,0].real,.02)
        self.assertAlmostEqual(transformed[2][0,0].real,2.15)
        self.assertAlmostEqual(transformed[3][0,0].real,2.30)
        self.assertAlmostEqual(base[0][0,0].real,2.25)

    def test_pair_labeled_todo_on_missing_both_polarizations(self):
        m={"classification":{"polarization_class":"mixed"},"status":"candidate_unverified",
           "residual_relative":1e-9,"neff_real":1.445}
        result=derive_modal_birefringence([m],1550)
        self.assertEqual(result["status"],"TODO_POLARIZATION_MODE_ID")
        self.assertIsNone(result["B_modal_candidate"])

    @unittest.skipUnless(HAVE_DEPS,"requires Gmsh + scikit-fem")
    def test_actual_panda_stress_rod_and_vector_eigen_candidate(self):
        from geometry import build_mesh
        from stress_optic import solve_thermal_stress
        from vector_fem import solve
        d=small_pm()
        mesh=build_mesh(d)
        self.assertIn("stress_rod",set(mesh["material_ids"]))
        stress,delta=solve_thermal_stress(d,mesh,fake_props(),max_samples=4)
        self.assertGreater(stress["mechanics"]["max_abs_stress_Pa"],0)
        self.assertIsNotNone(delta)
        self.assertGreater(float(np.max(np.abs(delta))),0)
        with_pm=dict(mesh,pm_eps_delta=delta,pm_source_sha256=stress["input_sha256"])
        out=solve(d,with_pm,n_modes=1,target_neff=1.446,max_sample_points=4)
        self.assertEqual(out["engine_version"],"0.8.6")
        self.assertEqual(out["photoelastic_coupling"]["enabled"],True)
        self.assertFalse(out["validation"]["loss_dB_km_certified"])
        self.assertGreater(len(out["modes"]),0)
        self.assertIn("neff_real",out["modes"][0])
        self.assertFalse(out["validation"].get("pm_modal_birefringence_certified",False))
        # This smoke does not compare against independently verified COMSOL modes.

if __name__=="__main__":unittest.main()

