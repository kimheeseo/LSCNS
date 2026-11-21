"""Synthetic acceptance-logic fixtures; no external physics certification."""
import hashlib,json,tempfile,unittest
import copy
import math
from pathlib import Path
from physics_gate import artifact,external_gate,evaluate,normalized_bent_csv,independent_convergence
from step4_v086 import digest


class PhysicsGate(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup);self.root=Path(self.temp.name)

    def save(self,name,data):
        text=data if isinstance(data,str) else json.dumps(data)
        (self.root/name).write_text(text)
        return {'path':name,'sha256':hashlib.sha256(text.encode()).hexdigest()}

    def package(self):
        field=self.save('field.json',{'tracking_fields':{'points_um':[[0,0]],'weights_um2':[1.],
                                                       'E':[[[1.,0.],[0.,0.],[0.,0.]]]}})
        modes=['mode_'+str(i) for i in range(10)]
        table='wavelength_nm,mode_id,neff_real,neff_imag\n'+''.join('1550,'+m+',1.4,1e-6\n' for m in modes)
        c=self.save('table.csv',table)
        mapping=self.save('identity.json',{'reviewed_by':'TEST FIXTURE','basis':'synthetic identity regression',
            'mode_mapping':[{'wavelength_nm':1550.,'reference_mode_id':m,'calculated_mode_id':m,
                             'reference_field':field,'calculated_field':field,'minimum_subspace_correlation':1.} for m in modes]})
        axes={}
        for axis in ('mesh','window','pml_thickness','pml_strength','pml_order'):
            rows=[]
            for i in range(3):
                d={'mesh':{'resolution':1},'domain_radius_um':10,'boundary':{'thickness_um':1,'strength':1,'order':1}}
                if axis=='mesh':d['mesh']['resolution']=i+1
                elif axis=='window':d['domain_radius_um']=i+10
                else:d['boundary'][{'pml_thickness':'thickness_um','pml_strength':'strength','pml_order':'order'}[axis]]=i+1
                modes=[]
                for pol in range(2):
                    e=[[[1.,0.],[0.,0.],[0.,0.]]] if pol==0 else [[[0.,0.],[1.,0.],[0.,0.]]]
                    modes.append({'neff_real':1.4,'neff_imag':1e-6,'status':'candidate_unverified',
                                  'tracking_fields':{'points_um':[[0,0]],'weights_um2':[1.],'E':e}})
                spec=self.save(axis+'_'+str(i)+'.json',{'input':d,'candidates':modes})
                rows.append({'level':i,'input_sha256':digest(d),'result':spec,'mode_indices':[0,1]})
            axes[axis]=rows
        conv=self.save('convergence.json',{'tolerances':{'Re_absolute':1e-5,'Im_relative':.05,'Im_absolute':1e-12},
            'axes':axes})
        return {'reference_csv':c,'calculated_csv':c,'exact_parameters':self.save('parameters.txt','synthetic fixture only'),
            'identity':mapping,'convergence':conv,
            'model_manifest':self.save('model.json',{'comsol_version':'TEST','source_url':'synthetic fixture only',
                                                   'input_equivalence_reviewed_by':'TEST'}),
            'reference_imag_sign':1,'calculated_imag_sign':1,
            'tolerances':{'Re_ppm':1.,'Im_relative':.05,'Im_absolute':1e-12}}

    def test_complete_synthetic_package_evaluates_and_missing_artifact_blocks(self):
        manifest=self.package();self.assertEqual(external_gate('V2',manifest,self.root)['status'],'PASS')
        manifest['exact_parameters']=None
        self.assertEqual(external_gate('V2',manifest,self.root)['status'],'NEEDS_EVIDENCE')

    def test_bend_package_requires_consistent_raw_normalization(self):
        m=self.package()
        raw='wavelength_nm,mode_id,raw_real,raw_imag,r_average_m,r_power_m\n'+''.join('1550,mode_'+str(i)+',1.4,1e-6,.003,.003\n' for i in range(10))
        m['raw_bent_csv']=self.save('raw_bent.csv',raw)
        m['normalization_manifest']=self.save('normalization.json',{'representation':'neff_geometry','length_unit':'m',
            'imag_sign':1,'comparison_radius_basis':'geometry'})
        m['formulation_review']=self.save('formulation.json',{'reviewed_by':'TEST','equivalent_comparison':True})
        self.assertEqual(external_gate('V3',m,self.root)['status'],'PASS')
        m['raw_bent_csv']=self.save('wrong_bent.csv',raw.replace('1e-6','2e-6'))
        self.assertEqual(external_gate('V3',m,self.root)['status'],'NEEDS_EVIDENCE')

    def test_declared_identity_cannot_override_wrong_actual_field(self):
        m=self.package();r=json.loads((self.root/'identity.json').read_text())
        field=self.save('wrong_field.json',{'tracking_fields':{'points_um':[[0,0]],'weights_um2':[1.],
                                                    'E':[[[0.,0.],[1.,0.],[0.,0.]]]}})
        for row in r['mode_mapping']:row['calculated_field']=field
        m['identity']=self.save('wrong_identity.json',r)
        self.assertEqual(external_gate('V2',m,self.root)['status'],'NEEDS_EVIDENCE')

    def test_null_tolerance_sign_or_field_mapping_never_passes(self):
        for key in ('tolerances','reference_imag_sign','identity'):
            m=self.package();m[key]={} if key=='tolerances' else None
            self.assertEqual(external_gate('V2',m,self.root)['status'],'NEEDS_EVIDENCE')

    def test_sha_tampering_rejected(self):
        m=self.package();(self.root/'table.csv').write_text('tampered')
        self.assertEqual(external_gate('V2',m,self.root)['status'],'NEEDS_EVIDENCE')
        with self.assertRaises(ValueError):artifact(self.root,{'path':'../escape','sha256':'0'*64})

    def test_residual_or_single_level_not_convergence(self):
        m=self.package();r=json.loads((self.root/'convergence.json').read_text())
        r['axes']['pml_strength']=r['axes']['pml_strength'][:1]
        with self.assertRaises(ValueError):independent_convergence(self.root,self.save('short.json',r))
        r=json.loads((self.root/'convergence.json').read_text())
        data=json.loads((self.root/'mesh_2.json').read_text())
        for m in data['candidates']:m['neff_imag']=2e-6
        r['axes']['mesh'][-1]['result']=self.save('mesh_2_bad.json',data)
        self.assertFalse(independent_convergence(self.root,self.save('bad.json',r)))

    def test_convergence_cannot_hide_coupled_input_changes(self):
        self.package();r=json.loads((self.root/'convergence.json').read_text())
        data=json.loads((self.root/'mesh_2.json').read_text());data['input']['domain_radius_um']=20
        r['axes']['mesh'][-1].update(result=self.save('coupled.json',data),input_sha256=digest(data['input']))
        with self.assertRaises(ValueError):independent_convergence(self.root,self.save('coupled_manifest.json',r))

    def test_bent_raw_radius_contract(self):
        raw='wavelength_nm,mode_id,raw_real,raw_imag,r_average_m,r_power_m\n1550,a,1.445,-2e-8,.003,.003004\n'
        norm={'representation':'neff_geometry','imag_sign':-1,'comparison_radius_basis':'power'}
        self.assertIn('1.443',normalized_bent_csv(raw,norm))
        with self.assertRaises(KeyError):normalized_bent_csv(raw.replace(',r_power_m',''),norm)

    def test_baseline_hash_and_missing_items_fail_closed(self):
        baseline={'items':[{'id':'V'+str(i),'status':'PASS','scope':'synthetic test'} for i in range(1,10)]}
        h=hashlib.sha256(json.dumps(baseline,sort_keys=True,separators=(',',':'),ensure_ascii=False).encode()).hexdigest()
        self.save('external_manifest.json',{'baseline_canonical_sha256':h})
        r=evaluate(self.root,baseline)
        self.assertFalse(r['all_required_passed']);self.assertEqual(r['items']['V2']['status'],'NEEDS_EVIDENCE')
        baseline['items'][0]['status']='FAIL'
        with self.assertRaises(ValueError):evaluate(self.root,baseline)

    def baseline(self):
        b={'items':[{'id':'V'+str(i),'status':'PASS','scope':'synthetic fixture'} for i in range(1,10)]}
        h=hashlib.sha256(json.dumps(b,sort_keys=True,separators=(',',':'),ensure_ascii=False).encode()).hexdigest()
        return b,{'baseline_canonical_sha256':h}

    def test_v6_computes_reference_error_instead_of_trusting_declared_zero(self):
        p=self.package();baseline,manifest=self.baseline();(self.root/'V6').mkdir()
        r={'cases':[{'wavelength_nm':1200.,'pitch_um':2.,'diameter_ratio':.6,
            'FSM':[{'n_FSM':1.3}]*3,'FM_mesh':[{'status':'TRACKED_CANDIDATE','neff_real':[1.4,1.4]}]*3}],
            'policy':{'FSM_absolute_neff_last_two_changes':1e-5,'FM_absolute_neff_last_two_changes':1e-5}}
        (self.root/'V6/summary.json').write_text(json.dumps(r))
        curve={'rows':[{'wavelength_nm':1200.,'pitch_um':2.,'diameter_ratio':.6,
                        'V_PCF':2*math.pi*2/1.2*math.sqrt(1.4**2-1.3**2),'uncertainty_absolute':0.}]}
        qualification={'reviewed_by':'TEST','cutoff_bracket':[1000,1200],
            'primary_reference_data':self.save('curve.json',curve),'HOM_field_evidence':self.save('HOM.json',{'synthetic_fixture':True}),
            'reference_V_relative_tolerance':.01,'reference_V_relative_error':0.}
        manifest['V6']={'identity':p['identity'],'convergence':p['convergence'],
                        'HOM_reference_comparison':self.save('qualification.json',qualification)}
        self.save('external_manifest.json',manifest)
        self.assertEqual(evaluate(self.root,baseline)['items']['V6']['status'],'PASS')
        curve['rows'][0]['V_PCF']=1e-6
        qualification['primary_reference_data']=self.save('wrong_curve.json',curve)
        manifest['V6']['HOM_reference_comparison']=self.save('wrong_qualification.json',qualification)
        self.save('external_manifest.json',manifest)
        self.assertEqual(evaluate(self.root,baseline)['items']['V6']['status'],'FAIL')

    def test_v7_raw_refinement_and_branch_gap_gate(self):
        p=self.package();baseline,manifest=self.baseline();(self.root/'V7_adaptive').mkdir()
        def rows(step):return [{'wavelength_nm':x,'status':'TRACKED_CANDIDATE','alpha_total_dB_m':y}
                               for x,y in [(1000-step,1),(1000,3),(1000+step,1)]]
        r={'cases':rows(100),'arrow_lambda_nm':1000.}
        (self.root/'V7_adaptive/summary.json').write_text(json.dumps(r))
        levels=[{'spectrum':self.save('spectrum_'+str(i)+'.json',{'cases':rows(step)})} for i,step in enumerate([100,50,25])]
        q={'levels':levels,'relative_tolerance':.01}
        manifest['V7']={'identity':p['identity'],'convergence':p['convergence'],'spectral_refinement':self.save('refinement.json',q)}
        self.save('external_manifest.json',manifest)
        self.assertEqual(evaluate(self.root,baseline)['items']['V7']['status'],'PASS')
        r['cases'][1]['status']='FAILED';(self.root/'V7_adaptive/summary.json').write_text(json.dumps(r))
        self.assertEqual(evaluate(self.root,baseline)['items']['V7']['status'],'FAIL')


if __name__=='__main__':unittest.main()
