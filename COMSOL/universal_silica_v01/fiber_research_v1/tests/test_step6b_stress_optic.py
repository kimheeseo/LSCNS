"""STEP6b analytical synthetic tests. Numbers are NOT actual PANDA rod material data."""
import copy
import math
import sys
import unittest
from pathlib import Path
import numpy as np
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from schema import load_presets,validate_design
from stress_optic import validate_pm_config,isotropic_photoelastic,solve_thermal_stress,ENGINE_VERSION
SOURCE="https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.stress_optical/stress_optical.html"

def design(variant="panda"):
    return copy.deepcopy(load_presets()["polarization_maintaining__"+variant])

def simple_mesh():
    # Synthetic 2x1 rectangle, two triangles core, two stress_rod and background
    x=np.array([[0,0],[1,0],[2,0],[0,1],[1,1],[2,1],[1,2],[2,2]],float)
    t=np.array([[0,1,3],[1,4,3],[1,2,4],[2,5,4],[4,5,6],[5,7,6]],int)
    return {"points_um":x,"triangles":t,
            "material_ids":np.array(["core","core","stress_rod","stress_rod","background","background"],dtype="U32")}

def config(dT=-500.,fill_photo=True,rod_factor=1.0):
    # Synthetic elastic/photoelastic values are intentionally artificial and
    # used only to verify algebra. The URL cites method, NOT these fake values.
    out={"schema":"LSCNS_step6b_materials_v1","delta_T_K":dT,"materials":{}}
    for k in ["core","background","stress_rod"]:
        m={"E_Pa":8.0e10,"poisson_ratio":.2,
           "alpha_1_per_K":5e-7*(rod_factor if k=="stress_rod" else 1),
           "mechanical_source_url":SOURCE}
        if fill_photo:
            m.update({"p11":.12,"p12":.27,"photoelastic_source_url":SOURCE})
        out["materials"][k]=m
    return out

class TestStep6b(unittest.TestCase):
    def test_engine(self):
        self.assertEqual(ENGINE_VERSION,"0.8.2")

    def test_homogeneous_free_generalized_plane_strain_zero_stress(self):
        d=design();mesh=simple_mesh()
        r,delta=solve_thermal_stress(d,mesh,config(dT=-500))
        self.assertAlmostEqual(r["mechanics"]["axial_strain_uniform"],-500*5e-7,delta=1e-10)
        self.assertLess(r["mechanics"]["max_abs_stress_Pa"],.5)
        self.assertLess(abs(r["mechanics"]["normalized_axial_resultant_Pa_area_fraction"]),.5)
        self.assertEqual(r["photoelastic"]["core_material_birefringence"]["status"],"MATERIAL_PROXY_ONLY")
        self.assertLess(abs(r["photoelastic"]["core_material_birefringence"]["B_core_area_weighted"]),1e-9)
        self.assertEqual(delta.shape,(len(mesh["triangles"]),4))
        self.assertTrue(np.isfinite(delta).all())
        self.assertEqual(len(r["input_sha256"]),64)

    def test_material_mismatch_changes_mechanics(self):
        d=design();mesh=simple_mesh()
        r,_=solve_thermal_stress(d,mesh,config(rod_factor=6.))
        self.assertGreater(r["mechanics"]["max_abs_stress_Pa"],1000)
        self.assertAlmostEqual(r["mechanics"]["delta_T_K"],-500)

    def test_zero_temperature(self):
        r,_=solve_thermal_stress(design(),simple_mesh(),config(dT=0))
        self.assertAlmostEqual(r["mechanics"]["max_abs_stress_Pa"],0,delta=.0001)

    def test_isotropic_strain_optic_sign_and_shear(self):
        Q,e,n=isotropic_photoelastic(1.45,.12,.27,1e-3,0,0,1e-4)
        self.assertAlmostEqual(Q[0,0]-Q[1,1],(.12-.27)*1e-3,places=12)
        self.assertAlmostEqual(Q[0,1],(.12-.27)*1e-4,places=12)
        self.assertAlmostEqual(Q[1,0],Q[0,1])
        self.assertGreater(n[0],n[1])
        self.assertGreater(e[0,0],0)
        with self.assertRaises(ValueError):
            isotropic_photoelastic(0,.12,.27,0,0,0,0)

    def test_missing_photoelastic_coeffs_gives_todo(self):
        r,eps=solve_thermal_stress(design(),simple_mesh(),config(fill_photo=False))
        self.assertIsNone(eps)
        self.assertEqual(r["photoelastic"]["core_material_birefringence"]["status"],"TODO_PHOTOELASTIC")
        self.assertGreater(r["mechanics"]["physical_triangles"],0)

    def test_missing_mechanics_source_rejected(self):
        m=config();del m["materials"]["stress_rod"]["mechanical_source_url"]
        with self.assertRaisesRegex(ValueError,"source"):solve_thermal_stress(design(),simple_mesh(),m)

    def test_missing_mechanical_value_rejected(self):
        m=config();m["materials"]["stress_rod"]["E_Pa"]=None
        with self.assertRaisesRegex(ValueError,"E_Pa"):solve_thermal_stress(design(),simple_mesh(),m)

    def test_missing_photo_source_rejected(self):
        m=config();del m["materials"]["core"]["photoelastic_source_url"]
        with self.assertRaisesRegex(ValueError,"source"):solve_thermal_stress(design(),simple_mesh(),m)

    def test_no_unsupported_panda_without_rod_label(self):
        mesh=simple_mesh()
        mesh["material_ids"][2]="background";mesh["material_ids"][3]="background"
        with self.assertRaisesRegex(ValueError,"stress_rod"):solve_thermal_stress(design(),mesh,config())

    def test_schema_requires_separate_rod_material(self):
        d=design()
        del d["materials"]["stress_rod"]
        with self.assertRaisesRegex(ValueError,"stress_rod"):validate_design(d)

    def test_pml_excluded(self):
        m=simple_mesh()
        # Extra synthetic PML triangle, no mechanics coefficient.
        m["triangles"]=np.vstack([m["triangles"],np.array([2,7,5])])
        m["material_ids"]=np.append(m["material_ids"],"pml")
        r,eps=solve_thermal_stress(design(),m,config())
        self.assertEqual(r["mechanics"]["physical_triangles"],6)
        self.assertTrue(np.all(eps[-1]==0))

    def test_input_change_hash_invalidation(self):
        a,_=solve_thermal_stress(design(),simple_mesh(),config(dT=-500))
        b,_=solve_thermal_stress(design(),simple_mesh(),config(dT=-501))
        self.assertNotEqual(a["input_sha256"],b["input_sha256"])

if __name__=="__main__":unittest.main()
