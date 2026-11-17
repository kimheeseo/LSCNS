"""STEP 6a v0.8.1 — analytically grounded bend unit tests (NOT COMSOL V3)."""
import copy
import math
import sys
import unittest
from pathlib import Path
import numpy as np
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from bending import (ENGINE_VERSION,configure_bend,conformal_factor,bend_per_turn,
                     bend_radius_direction_sweep,COMSOL_V3_SOURCE)
from schema import load_presets
from vector_fem import _tensor

SOURCE="https://optiwave.wiki/kb/optical_fiber_bending_loss_with_stress_effect/"

def design():
    return copy.deepcopy(load_presets()["index_guiding__step"])


def mock_mode(neff,alpha,which):
    q={"points_um":[[0,0],[1,0],[1,1],[0,1]],
       "weights_um2":[1,1,1,1],
       "Ex":[[1,0]]*4 if which==0 else [[0,0]]*4,
       "Ey":[[0,0]]*4 if which==0 else [[1,0]]*4}
    return {"neff_real":neff,"neff_imag":1e-9,
            "alpha_total_dB_m":alpha,"status":"candidate_unverified",
            "field_quadrature":q}


def fake_solver(case,mesh,n_modes,target_neff):
    # Deterministic symbolic fixture, NOT an optical model. For mode identity testing only.
    b=case.get("_bend")
    bend_increment=0 if b is None else .01/b["radius_mm"]
    modes=[mock_mode(1.446+bend_increment,.1+bend_increment,0),
           mock_mode(1.445+bend_increment,.2+bend_increment,1)]
    if b and b["radius_mm"]<8:
        modes=modes[::-1]
    return {"modes":modes}


class TestBendFormula(unittest.TestCase):
    def test_version(self):
        self.assertEqual(ENGINE_VERSION,"0.8.1")

    def test_conformal_units_zero_and_direction(self):
        R=10.0
        self.assertAlmostEqual(float(conformal_factor(100.,0.,R,0)),1.01)
        self.assertAlmostEqual(float(conformal_factor(-100.,0.,R,0)),.99)
        self.assertAlmostEqual(float(conformal_factor(100.,0.,R,90)),1.0,places=12)
        self.assertAlmostEqual(float(conformal_factor(0.,100.,R,90)),1.01)
        self.assertAlmostEqual(float(conformal_factor(100.,100.,R,180)),.99)
        self.assertAlmostEqual(float(conformal_factor(100.,0.,R,360)),1.01)

    def test_conformal_tensor_epsilon_square(self):
        d=configure_bend(design(),50.,0)
        x=np.array([[100.,-100.],[0.,0.]])
        e=np.ones((2,2),dtype=complex)*2.25
        ex,*_= _tensor(x,e,d)
        self.assertAlmostEqual(ex[0,0].real,2.25*(1+100./50000)**2,places=10)
        self.assertAlmostEqual(ex[0,1].real,2.25*(1-100./50000)**2,places=10)
        straight=copy.deepcopy(d)
        del straight["_bend"]
        ex0,*_=_tensor(x,e,straight)
        self.assertAlmostEqual(ex0[0,0].real,2.25,places=11)

    def test_radius_edge_and_invalid(self):
        d=design()
        extent=d["domain_radius_um"]+d["boundary"]["thickness_um"]
        with self.assertRaises(ValueError):configure_bend(d,extent/1000)
        with self.assertRaises(ValueError):configure_bend(d,0)
        with self.assertRaises(ValueError):configure_bend(d,-5)
        with self.assertRaises(ValueError):configure_bend(d,float("nan"))
        with self.assertRaises(ValueError):conformal_factor(-20000,0,10,0)
        with self.assertRaises(ValueError):conformal_factor(1,2,float("inf"),0)

    def test_stress_literature_is_explicit(self):
        d=design()
        with self.assertRaises(ValueError):
            configure_bend(d,10,0,reff_multiplier=1.2475)
        with self.assertRaises(ValueError):
            configure_bend(d,10,0,reff_multiplier=1.2475,stress_source_url=SOURCE)
        b=configure_bend(d,10,30,reff_multiplier=1.2475,
             stress_source_url=SOURCE,stress_source_note="SMF silica, source at 1550 nm")
        self.assertAlmostEqual(b["_bend"]["R_eff_mm"],12.475)
        self.assertEqual(b["_bend"]["stress_correction"]["source"],SOURCE)
        self.assertNotIn("_bend",d)
        with self.assertRaises(ValueError):
            configure_bend(d,10,reff_multiplier=-1,stress_source_url=SOURCE,stress_source_note="silica")
        with self.assertRaises(ValueError):
            configure_bend(d,10,stress_source_url=SOURCE)

    def test_per_turn_physical_R_and_signed(self):
        self.assertAlmostEqual(bend_per_turn(.1,10),.1*2*math.pi*.01)
        self.assertAlmostEqual(bend_per_turn(-.1,10),-.1*2*math.pi*.01)
        with self.assertRaises(ValueError):bend_per_turn(.1,0)

    def test_roundtrip_direction_and_ratio(self):
        b=configure_bend(design(),20.,-90.)
        self.assertAlmostEqual(b["_bend"]["direction_deg"],270.)
        self.assertAlmostEqual(b["_bend"]["outer_extent_over_R_eff"],
                               (b["domain_radius_um"]+b["boundary"]["thickness_um"])/20000.)

    def test_sweep_mode_swap_and_direction(self):
        r=bend_radius_direction_sweep(
            design(),[5.,10.],[0.,90.],mesh={"dummy":True},
            solve_fn=fake_solver,minimum_overlap=.75)
        self.assertEqual(r["engine_version"],"0.8.1")
        self.assertEqual(r["validation"]["V3"],"EVIDENCE_REQUIRED")
        self.assertEqual(len(r["cases"]),4)
        self.assertTrue(all(c["status"]=="COMPUTED_UNVERIFIED" for c in r["cases"]))
        self.assertEqual(r["radii_mm_descending"],[10.,5.])
        self.assertEqual(r["directions_deg"],[0.,90.])
        self.assertTrue(all(c["overlap_power"]>.99 for c in r["cases"]))
        self.assertAlmostEqual(r["cases"][1]["delta_total_dB_m"],.01/5)
        self.assertEqual(len(r["input_sha256"]),64)
        self.assertFalse(r["validation"]["loss_certified"])

    def test_input_hash_detects_radius_angle_source_changes(self):
        x=bend_radius_direction_sweep(design(),[5.],[0.],mesh={},
             solve_fn=fake_solver)
        y=bend_radius_direction_sweep(design(),[6.],[0.],mesh={},
             solve_fn=fake_solver)
        z=bend_radius_direction_sweep(design(),[5.],[90.],mesh={},
             solve_fn=fake_solver)
        w=bend_radius_direction_sweep(design(),[5.],[0.],mesh={},
             solve_fn=fake_solver,reff_multiplier=1.2475,stress_source_url=SOURCE,
             stress_source_note="SMF silica paper, 1550 nm")
        self.assertEqual(len({p["input_sha256"] for p in (x,y,z,w)}),4)

    def test_no_fabricated_loss_after_overlap_mismatch(self):
        def broken_solver(case,mesh,n_modes,target_neff):
            r=fake_solver(case,mesh,n_modes,target_neff)
            if "_bend" in case:
                for m in r["modes"]:
                    m["field_quadrature"]["points_um"]=[[0,0],[2,0],[2,2],[0,2]]
            return r
        r=bend_radius_direction_sweep(design(),[10],[0],mesh={},
             solve_fn=broken_solver)
        c=r["cases"][0]
        self.assertEqual(c["status"],"SOLVE_FAILED")
        self.assertIsNone(c["alpha_total_dB_m"])
        self.assertIsNone(c["delta_total_dB_m"])

    def test_requires_pml_and_no_duplicate_radius(self):
        d=design();d["boundary"]["type"]="pec"
        with self.assertRaises(ValueError):
            bend_radius_direction_sweep(d,[10],[0],mesh={},solve_fn=fake_solver)
        with self.assertRaises(ValueError):
            bend_radius_direction_sweep(design(),[10,10],[0],mesh={},solve_fn=fake_solver)

    def test_reference_is_not_unjustified_comsol_numeric(self):
        self.assertIn("step_index_fiber_bend",COMSOL_V3_SOURCE)
        r=bend_radius_direction_sweep(design(),[10],[0],mesh={},solve_fn=fake_solver)
        self.assertNotIn("benchmark_neff",r["validation"])

if __name__=="__main__":unittest.main()
