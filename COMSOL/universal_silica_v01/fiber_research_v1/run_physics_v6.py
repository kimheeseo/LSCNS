"""FSM cutoff and field-matched finite-PCF convergence evidence, not ESM proof."""
import argparse,copy
from pathlib import Path
import numpy as np
from schema import load_presets
from modal_diagnostics import disk_probes
from modal_diagnostics import weighted_columns
from step6_fields_v088 import common,unpack
from validation_photonics_v089 import fsm_index
from modal_search_evidence import search,choose_pair,write_json,source_hashes


def convergence(values,tolerance):
    values=np.asarray(values,float)
    if len(values)<3 or not np.isfinite(values).all():return False
    return bool(np.all(np.abs(np.diff(values)[-2:])<=tolerance))


def relative_mode_candidates(pool,fm):
    """Field projection separates FM-subspace matches from possible HOMs; labels are provisional."""
    result=[]
    for m in pool:
        _,w,f=common(fm+[m]);a=weighted_columns(f[:,:,:len(fm)],w);b=weighted_columns(f[:,:,-1:],w)[:,0]
        u,s,_=np.linalg.svd(a,full_matrices=False);u=u[:,s>s[0]*1e-8]
        projection=float(np.clip(np.linalg.norm(u.conj().T@b)/np.linalg.norm(b),0,1))
        label='FM_SUBSPACE_MATCH' if projection>=.8 else 'HOM_CANDIDATE_NOT_IDENTIFIED' if m['neff_real']<min(v['neff_real'] for v in fm) else 'OTHER_CORE_CANDIDATE'
        result.append({'candidate_id':m['candidate_id'],'FM_subspace_projection':projection,'label':label,
                       'neff_real':m['neff_real'],'neff_imag':m['neff_imag'],
                       'core_disk_Et_fraction':m['core_disk_Et_fraction'],
                       'field_classification':m.get('classification',{})})
    return result


def base_design(lam):
    d=copy.deepcopy(load_presets()['solid_pcf__hex'])
    d.update(name='V6 field-matched PCF convergence evidence',wavelength_nm=lam,domain_radius_um=8.)
    d['parameters'].update(pitch_um=2.,d_over_pitch=.6,rings=3,defect_cells=1)
    d['materials']['background']={'model':'constant','n':1.444}
    d['materials']['air']={'model':'constant','n':1.}
    d['mesh'].update(elements_per_wavelength=4,through_wall_elements=3,bulk_max_um=.8,curvature_segments=64)
    d['boundary'].update(type='pml',thickness_um=1.5,strength=3.,order=3)
    d['solver']={'max_dofs':180000,'max_iterations':700,'tolerance':1e-9}
    return d


def run(out):
    out=Path(out);cases=[]
    # Predeclared numerical screening policy, not a tolerance attributed to the paper.
    policy={'FSM_absolute_neff_last_two_changes':1e-5,'FM_absolute_neff_last_two_changes':1e-5,
            'minimum_subspace_correlation':.8,'basis':'engineering diagnostic policy declared before run'}
    p,w=disk_probes(7.44,1.4,radial_order=8,angular_points=16)
    first_fm=None
    for lam in (1200.,1800.):
        row={'wavelength_nm':lam,'pitch_um':2.,'diameter_ratio':.6,'FSM':[],'FM_mesh':[]}
        for cutoff in (2,3,4,5,6):
            r=fsm_index(2.,lam,.6,cutoff);row['FSM'].append(r)
            write_json(out/'progress.json',{'cases':cases,'current':row})
            print('FSM',lam,cutoff,r['n_FSM'],flush=True)
        row['FSM_converged']=convergence([r['n_FSM'] for r in row['FSM']],1e-5)
        d=base_design(lam);prev=None
        prediction=1.425 if first_fm is None else float(np.sqrt(1.444**2-(lam/1200.)**2*(1.444**2-first_fm**2)))
        for j,(epw,curvature,bulk) in enumerate(((4,64,.8),(5,96,.6),(6,128,.45))):
            dd=copy.deepcopy(d);dd['mesh'].update(elements_per_wavelength=epw,curvature_segments=curvature,bulk_max_um=bulk)
            targets=[prediction,1.444-.03*(lam/1200.)**2,prediction-.015,prediction+.015]
            if prev is not None:targets.insert(0,float(np.mean([m['neff_real'] for m in prev])))
            label=f'lambda_{int(lam)}_mesh_{j}'
            pool,record=search(dd,targets,p,w,1.4,out/label,previous=prev,stop_on_pair=True)
            pair,selection=choose_pair(pool,prev);record.update(selection)
            record['candidate_file']=label+'/candidates.json'
            row['FM_mesh'].append(record)
            if pair is None:break
            prev=pair
            if lam==1200. and j==0:first_fm=float(np.mean([m['neff_real'] for m in pair]))
        fm=[float(np.mean(r['neff_real'])) for r in row['FM_mesh'] if r['status']=='TRACKED_CANDIDATE']
        row['FM_mesh_converged']=len(fm)==3 and convergence(fm,1e-5)
        if len(fm)==3:
            dd=copy.deepcopy(d);dd.update(domain_radius_um=10.);dd['parameters']['rings']=4
            label=f'lambda_{int(lam)}_rings_4'
            pool,record=search(dd,[fm[-1]],p,w,1.4,out/label,previous=prev,stop_on_pair=True)
            pair,selection=choose_pair(pool,prev);record.update(selection,candidate_file=label+'/candidates.json')
            row['ring_window_probe']=record
        if fm and prev is not None:
            label=f'lambda_{int(lam)}_HOM_search'
            pool,record=search(d,[fm[-1]-.015,fm[-1]-.035],p,w,1.4,out/label)
            record['candidate_file']=label+'/candidates.json'
            record['relative_to_last_FM']=relative_mode_candidates(pool,prev)
            row['HOM_search']=record
        nf=row['FSM'][-1]['n_FSM'];rad=fm[-1]**2-nf**2 if fm else None
        row['V_PCF_candidate']=float(2*np.pi*2/(lam*.001)*np.sqrt(rad)) if rad is not None and rad>0 else None
        row['HOM_identity']='NEEDS_EVIDENCE';row['primary_reference_quantitative_comparison']=None
        row['status']='NEEDS_EVIDENCE';cases.append(row)
        write_json(out/'progress.json',{'cases':cases})
    result={'schema':'LSCNS_physics_V6_convergence_v1','cases':cases,'policy':policy,
            'source_sha256':source_hashes(),'status':'NEEDS_EVIDENCE','ESM_certified':False,
            'missing':['independent FM/HOM identification and HOM cutoff','independent PML/window/ring convergence',
                       'quantitative primary-reference curve and uncertainty','broader wavelength and d/Lambda boundary assessment'],
            'scope':'d/Lambda=.6 finite PCF samples; not an ESM certification near .406'}
    write_json(out/'summary.json',result);return result


if __name__=='__main__':
    a=argparse.ArgumentParser();a.add_argument('--out',required=True);args=a.parse_args();print(run(args.out)['status'])
