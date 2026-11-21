"""Reproducible multi-shift ARF diagnostic scan with bounded adaptive bisection."""
import argparse,copy,json
from pathlib import Path
import numpy as np
from schema import load_presets
from modal_diagnostics import disk_probes
from modal_search_evidence import search,choose_pair,peak_evidence,write_json,source_hashes


def design():
    d=copy.deepcopy(load_presets()['anti_resonant__single_ring'])
    d.update(name='V7 evidence recovery synthetic single-ring ARF',domain_radius_um=13.)
    d['parameters'].update(core_radius_um=4.,tube_radius_um=3.5,wall_thickness_um=.55,tube_count=6,gap_um=.4)
    for name in ('background','air'):d['materials'][name]={'model':'constant','n':1.}
    d['materials']['silica']={'model':'constant','n':1.45}
    d['mesh'].update(elements_per_wavelength=4,through_wall_elements=3,curvature_segments=64,bulk_max_um=.9)
    d['boundary'].update(thickness_um=1.5,strength=3.,order=3)
    d['solver']={'max_dofs':150000,'max_iterations':600,'tolerance':1e-9}
    return d


def run(out,max_bisections=1):
    out=Path(out);d=design();arrow=1155.;rows=[];adaptive=[];previous=None;previous_lam=None
    p,w=disk_probes(.93*13,4.,radial_order=8,angular_points=16)
    def attempt(lam,prev,label):
        dd=copy.deepcopy(d);dd['wavelength_nm']=float(lam)
        seed=float(np.sqrt(1-(2.4048255577*(lam*.001)/(2*np.pi*4.4))**2))
        targets=[seed,seed-.003,seed+.003]
        if prev is not None:targets.insert(0,float(np.mean([m['neff_real'] for m in prev])))
        pool,record=search(dd,targets,p,w,4.,out/label,previous=prev,stop_on_pair=True,cache_root=out)
        pair,selection=choose_pair(pool,prev);record.update(selection,wavelength_nm=float(lam),
            seed_basis='empty-capillary shift only, not a reference answer',branch_id='core_pair_0' if pair else None)
        record['candidate_file']=str(Path(label)/'candidates.json')
        return pair,record
    def advance(lam,prev,start,depth,label):
        pair,row=attempt(lam,prev,label)
        if pair is None and prev is not None and depth>0 and start is not None:
            adaptive.append(row)
            mid=(start+lam)/2
            mp,mr=advance(mid,prev,start,depth-1,label+'_mid');adaptive.append(mr)
            if mp is not None:
                return advance(lam,mp,mid,depth-1,label+'_retry')
        return pair,row
    if (out/'progress.json').exists():
        prior=json.loads((out/'progress.json').read_text());rows=prior['cases'];adaptive=prior['adaptive_attempts']
        for row in rows:
            if row['status']=='TRACKED_CANDIDATE':
                data=json.loads((out/row['candidate_file']).read_text())
                expected=copy.deepcopy(d);expected['wavelength_nm']=row['wavelength_nm']
                from step4_v086 import digest
                if data['input_sha256']!=digest(expected):raise ValueError('resume design mismatch')
                previous=[data['candidates'][i] for i in row['candidate_ids']];previous_lam=row['wavelength_nm']
    for j,lam in enumerate(arrow*np.array([.85,.9,.95,1.,1.05,1.1,1.15])):
        if j<len(rows):continue
        pair,row=advance(float(lam),previous,previous_lam,max_bisections,f'lambda_{j}')
        rows.append(row)
        write_json(out/'progress.json',{'cases':rows,'adaptive_attempts':adaptive})
        # Never silently restart after a gap: any later recovery must overlap the last accepted pair.
        if pair is not None:previous=pair;previous_lam=float(lam)
    peak=peak_evidence(rows,arrow)
    result={'schema':'LSCNS_physics_V7_recovery_v1','cases':rows,'adaptive_attempts':adaptive,
            'arrow_lambda_nm':arrow,'peak':peak,'status':peak['status'],
            'source_sha256':source_hashes(),'max_bisections':max_bisections,
            'modal_identity_certified':False,'mesh_PML_window_converged':False,'loss_certified':False,
            'scope':'Et field screening and branch diagnostics; no independent FM certification'}
    write_json(out/'summary.json',result);return result


if __name__=='__main__':
    a=argparse.ArgumentParser();a.add_argument('--out',required=True);a.add_argument('--max-bisections',type=int,default=1)
    args=a.parse_args();print(run(args.out,args.max_bisections)['status'])
