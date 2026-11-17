"""STEP4 integration: result gating, fixed-quadrature matching and dispersion CLI."""
import copy
import sys
import unittest
from pathlib import Path
import numpy as np
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from step4_cli import postprocess_result,track_wavelength_sweep


def result(wavelength_nm,mode_order=False):
    d={"schema_version":"1.0.0","family":"index_guiding","variant":"step",
       "name":"synthetic TEST ONLY","wavelength_nm":float(wavelength_nm),
       "parameters":{"core_radius_um":4.1,"delta_n":.005},
       "materials":{"background":{"model":"pure_silica"},"core":{"model":"constant","n":1.45}}}
    pts=[[0,0],[1,0],[1,1],[0,1]]
    weights=[1,1,1,1]
    idx=float(wavelength_nm)-1550
    out=[]
    for i in range(2):
        ex=[[1,0]]*4 if i==0 else [[0,0]]*4
        ey=[[0,0]]*4 if i==0 else [[1,0]]*4
        q={"points_um":pts,"weights_um2":weights,"Ex":ex,"Ey":ey}
        neff=1.46-1.7e-9*idx*idx-i*.005
        out.append({"neff_real":neff,"neff_imag":.0,"classification":{
             "Aeff_um2":100.,"MFD_2nd_um":10.,"Aeff_basis":"Et=Ex,Ey only"},
             "field_quadrature":q})
    if mode_order:out=out[::-1]
    return {"schema":"LSCNS_fiber_research_vector_v1","input":d,"modes":out}


class CliTests(unittest.TestCase):
    def test_result_physics_status_and_optional_n2(self):
        data=result(1550.)
        o=postprocess_result(data)
        self.assertEqual(o["engine_version"],"0.8.0")
        self.assertEqual(o["modes"][0]["Aeff_Et_um2"],100.)
        self.assertIsNone(o["modes"][0]["gamma_W_inv_m_inv"])
        self.assertEqual(o["dispersion"]["status"],"TODO_TRACKED_SWEEP")
        o=postprocess_result(data,n2_m2_W=2.2e-20)
        self.assertGreater(o["modes"][0]["gamma_W_inv_m_inv"],0)

    def test_sweep_tracks_permuted_order(self):
        lams=np.arange(1510.,1600.,10.)
        raw={"schema":"LSCNS_step4_raw_sweep_v1",
             "runs":[result(l,mode_order=i%2==1) for i,l in enumerate(lams)]}
        ans=track_wavelength_sweep(raw)
        self.assertEqual(len(ans["curves"]),2)
        self.assertTrue(all(len(c["D_ps_nm_km"])==len(lams) for c in ans["curves"]))
        self.assertTrue(all(min(c["overlap_power"])>.99 for c in ans["curves"]))

    def test_reject_inconsistent_geometry_or_missing_full_field(self):
        data={"schema":"LSCNS_step4_raw_sweep_v1","runs":[result(l) for l in np.arange(1510.,1600.,10.)]}
        data["runs"][2]["input"]["parameters"]["core_radius_um"]=5
        with self.assertRaises(ValueError):track_wavelength_sweep(data)
        data["runs"][2]["input"]["parameters"]["core_radius_um"]=4.1
        del data["runs"][3]["modes"][0]["field_quadrature"]
        with self.assertRaises(ValueError):track_wavelength_sweep(data)

    def test_no_glass_n2_without_values(self):
        d=result(1550.)
        d["input"]["family"]="anti_resonant"
        o=postprocess_result(d,n2_m2_W=1e-20)
        self.assertIsNone(o["modes"][0]["gamma_W_inv_m_inv"])
        self.assertIsNotNone(o["arrow_overlay"])


if __name__=="__main__":unittest.main()
