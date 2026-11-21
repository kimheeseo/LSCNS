import unittest
import numpy as np
import tempfile,json
from pathlib import Path
from unittest.mock import patch
from modal_search_evidence import peak_evidence,choose_pair,search,write_json,source_hashes
from step4_v086 import digest


class ModalEvidence(unittest.TestCase):
    def test_field_bundle_is_lossless_and_hash_checked(self):
        from package_physics_evidence import package,unpack
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);p=root/'trial/candidates.json'
            raw='{"neff_imag":1.2345678901234567e-08}\n'
            p.parent.mkdir();p.write_text(raw);package(root);p.unlink()
            self.assertEqual(unpack(root),1);self.assertEqual(p.read_text(),raw)
            with (root/'candidates.zip').open('ab') as f:f.write(b'corruption')
            with self.assertRaises(ValueError):unpack(root)

    def mode(self, polarization, candidate_id, outer=False):
        e=np.zeros((4,3),complex)
        e[2 if outer else 0,polarization]=1
        return {'candidate_id':candidate_id,'neff_real':.99,'neff_imag':1e-5,
                'alpha_total_dB_m':1.,'core_disk_Et_fraction':.9,
                'tracking_fields':{'points_um':[[0,0],[.1,0],[3,0],[4,0]],
                                   'weights_um2':[1,1,1,1],
                                   'E':np.stack((e.real,e.imag),axis=-1).tolist()}}

    def test_polarization_rotation_and_order_do_not_break_subspace(self):
        a,b=self.mode(0,0),self.mode(1,1)
        pair,r=choose_pair([b,a],[a,b])
        self.assertIsNotNone(pair)
        self.assertAlmostEqual(r['minimum_subspace_correlation'],1.)

    def test_same_neff_wrong_spatial_branch_is_rejected(self):
        a,b=self.mode(0,0),self.mode(1,1)
        self.assertIsNone(choose_pair([self.mode(0,2,True),self.mode(1,3,True)],[a,b])[0])
        self.assertIsNone(choose_pair([a,a],[a,b])[0])

    def test_cached_actual_inputs_and_probes_are_required(self):
        from step6_fields_v088 import unpack
        a,b=self.mode(0,0),self.mode(1,1)
        for m in (a,b):m.update(status='candidate_unverified',search_target=.99)
        d={'wavelength_nm':1550};p,w,_=unpack(a)
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);saved=root/'old/candidates.json'
            data={'input':d,'input_sha256':digest(d),'n_modes_per_target':6,
                  'source_sha256':source_hashes(),'attempts':[{'target_neff':.99,'status':'COMPUTED','solver':{},'mesh_quality':{}}],
                  'candidates':[a,b]}
            write_json(saved,data)
            with patch('modal_search_evidence.build_mesh',side_effect=AssertionError('unexpected fresh solve')):
                pool,r=search(d,[.99],p,w,1.,root/'new',cache_root=root,stop_on_pair=True)
                self.assertEqual(len(pool),2);self.assertIn('cached_from',r['attempts'][0])
                origin=r['attempts'][0]['cached_from'];saved.write_text('overwritten later')
                import hashlib
                self.assertEqual(hashlib.sha256((root/origin['path']).read_bytes()).hexdigest(),origin['sha256'])
                with self.assertRaises(AssertionError):
                    saved.unlink()
                    search({'wavelength_nm':1560},[.99],p,w,1.,root/'changed',cache_root=root)

    def test_missing_branch_never_interpolated_to_peak(self):
        r=[{'wavelength_nm':x,'status':'TRACKED_CANDIDATE','alpha_total_dB_m':y} for x,y in [(900,1),(1000,3),(1100,1)]]
        self.assertEqual(peak_evidence(r,1000)['status'],'NEEDS_EVIDENCE')
        r[1]['status']='FAILED'
        self.assertEqual(peak_evidence(r,1000)['status'],'FAIL')

    def test_endpoint_or_invalid_spectrum_never_passes(self):
        for ys in ([3,2,1],[1,2,3],[1,-2,1],[1,float('nan'),1]):
            r=[{'wavelength_nm':x,'status':'TRACKED_CANDIDATE','alpha_total_dB_m':y} for x,y in zip([900,1000,1100],ys)]
            self.assertEqual(peak_evidence(r,1000)['status'],'FAIL')
        self.assertIsNone(choose_pair([])[0])


if __name__=='__main__':unittest.main()
