"""Auditable multi-shift modal search on common physical FE probes.

This is a diagnostic search, not independent identification or certification.
Rejected modes, fields, inputs, errors and solver/mesh metadata are retained.
"""
import copy
import hashlib
import itertools
import json
import platform
from pathlib import Path
from datetime import datetime, timezone
import numpy as np
from geometry import build_mesh
from vector_fem import solve
from step4_v086 import digest
from step6_fields_v088 import unpack, subspace_overlap, spatial_overlap


def write_json(path, data):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(data,ensure_ascii=False,allow_nan=False,separators=(',',':'))+'\n')


def source_hashes():
    root=Path(__file__).parent
    return {p:hashlib.sha256((root/p).read_bytes()).hexdigest() for p in
            ('modal_search_evidence.py','vector_fem.py','vector_solver_v085.py',
             'vector_spectrum.py','geometry.py','materials.py')}


def core_fraction(mode, radius):
    p,w,e=unpack(mode); intensity=np.sum(abs(e[:,:2])**2,axis=1)
    total=float(intensity@w)
    if total<=0: raise ValueError('zero Et probe norm')
    return float((intensity*(np.linalg.norm(p,axis=1)<radius))@w/total)


def search(design, targets, probes, weights, core_radius, out, n_modes=6, previous=None, stop_on_pair=False, cache_root=None):
    out=Path(out);out.mkdir(parents=True,exist_ok=True)
    mesh=None; attempts=[]; pool=[]; all_modes=[];cache={}
    if cache_root is not None:
        current=source_hashes()
        for path in sorted(Path(cache_root).rglob('candidates.json')):
            old=json.loads(path.read_text())
            if old.get('input_sha256')!=digest(design) or digest(old['input'])!=digest(design) or old.get('n_modes_per_target')!=n_modes:continue
            # Search logic can evolve; only unchanged actual solver/geometry/material code is reusable.
            if any(old.get('source_sha256',{}).get(name)!=value for name,value in current.items() if name!='modal_search_evidence.py'):continue
            for a in old['attempts']:
                modes=[m for m in old['candidates'] if m['search_target']==a['target_neff']]
                if a['status']=='COMPUTED' and modes and all(np.array_equal(unpack(m)[0],probes) and np.array_equal(unpack(m)[1],weights) for m in modes):
                    raw=path.read_bytes();sha=hashlib.sha256(raw).hexdigest()
                    snapshot=Path(cache_root)/'_cache_sources'/str(sha+'.json')
                    snapshot.parent.mkdir(parents=True,exist_ok=True)
                    if not snapshot.exists():snapshot.write_bytes(raw)
                    cache[a['target_neff']]=(modes,a,{'path':str(snapshot.relative_to(cache_root)),'sha256':sha})
    for k,target in enumerate(dict.fromkeys(float(x) for x in targets)):
        a={'target_neff':target,'candidate_indices':[]}
        try:
            if target in cache:
                modes,old,origin=cache[target]
                r={'modes':modes,'solver_diagnostics':old['solver'],'mesh_quality':old['mesh_quality']}
                a['cached_from']=origin
            else:
                if mesh is None:mesh=build_mesh(design)
                r=solve(design,mesh=mesh,n_modes=n_modes,target_neff=target,
                        probe_points=probes,probe_weights=weights,max_sample_points=6)
            a.update(solver=r['solver_diagnostics'],mesh_quality=r['mesh_quality'],status='COMPUTED')
            for m in r['modes']:
                m=copy.deepcopy(m);m['candidate_id']=len(all_modes);m['search_target']=target
                m['core_disk_Et_fraction']=core_fraction(m,core_radius)
                reasons=[]
                if m['status']!='candidate_unverified': reasons.append('SOLVER_SCREEN_REJECTED')
                if m['core_disk_Et_fraction']<.25: reasons.append('LOW_CORE_ET_PROXY')
                # Only collinear fields are duplicates; nearby neff alone never removes a mode.
                duplicate=None
                if not reasons:
                    for old in pool:
                        corr=subspace_overlap([old],[m])
                        if corr and corr[0]>1-1e-7:
                            duplicate=old['candidate_id'];break
                    if duplicate is not None: reasons.append('DUPLICATE_FIELD')
                m['rejection_reasons']=reasons;m['duplicate_of']=duplicate
                a['candidate_indices'].append(m['candidate_id']);all_modes.append(m)
                if not reasons:pool.append(m)
        except (RuntimeError,ValueError,MemoryError) as exc:
            a.update(status='FAILED',error=str(exc))
        attempts.append(a)
        print(out.name,'target',target,a['status'],'core candidates',len(pool),flush=True)
        # Expand the search only when field-based screening cannot recover the pair.
        # The executed target list remains in the evidence; no omitted solve is claimed.
        if stop_on_pair and choose_pair(pool,previous)[0] is not None:break
    data={'schema':'LSCNS_modal_candidate_evidence_v1','input':design,'input_sha256':digest(design),
          'timestamp_utc':datetime.now(timezone.utc).isoformat(),'python':platform.python_version(),
          'source_sha256':source_hashes(),'n_modes_per_target':n_modes,
          'stop_on_pair':stop_on_pair,'attempts':attempts,'candidates':all_modes,
          'probe_basis':'fixed physical-domain E probes, Et-intensity proxy; not power identity or precision integration',
          'physical_certified':False}
    write_json(out/'candidates.json',data)
    return pool,{'input_sha256':data['input_sha256'],'attempts':attempts,
                 'candidate_file':str(out/'candidates.json'),'candidate_count':len(all_modes)}


def choose_pair(pool, previous=None, minimum_correlation=.8, minimum_core=.25):
    groups=[]
    for a,b in itertools.combinations(pool,2):
        if abs(a['neff_real']-b['neff_real'])>1e-4:continue
        if min(a['core_disk_Et_fraction'],b['core_disk_Et_fraction'])<minimum_core:continue
        # Disallow a duplicate masquerading as a polarization pair.
        collinear=subspace_overlap([a],[b])
        if not collinear or collinear[0]>.98:continue
        if spatial_overlap(a,b)<.85:continue
        correlations=subspace_overlap(previous,[a,b]) if previous is not None else [1.,1.]
        if len(correlations)!=2 or min(correlations)<minimum_correlation:continue
        groups.append((min(correlations)*(a['core_disk_Et_fraction']+b['core_disk_Et_fraction'])/2,
                       [a,b],correlations))
    if not groups:return None,{'status':'FAILED','reason':'NO_CONTINUOUS_CORE_POLARIZATION_SUBSPACE'}
    _,pair,corr=max(groups,key=lambda v:v[0])
    return pair,{'status':'TRACKED_CANDIDATE','candidate_ids':[m['candidate_id'] for m in pair],
                 'minimum_subspace_correlation':min(corr),
                 'minimum_core_disk_Et_fraction':min(m['core_disk_Et_fraction'] for m in pair),
                 'neff_real':[m['neff_real'] for m in pair],
                 'neff_imag':[m['neff_imag'] for m in pair],
                 'alpha_total_dB_m':float(np.mean([m['alpha_total_dB_m'] for m in pair]))}


def peak_evidence(rows, arrow_nm):
    if len(rows)<3 or any(r.get('status')!='TRACKED_CANDIDATE' for r in rows):
        return {'status':'FAIL','reason':'MISSING_OR_BROKEN_BRANCH','observed_interior_peak_nm':None}
    wavelengths=np.array([r['wavelength_nm'] for r in rows],float)
    loss=np.array([r['alpha_total_dB_m'] for r in rows],float)
    if not np.isfinite(wavelengths).all() or not np.isfinite(loss).all() or np.any(np.diff(wavelengths)<=0) or np.any(loss<=0):
        return {'status':'FAIL','reason':'INVALID_SPECTRUM','observed_interior_peak_nm':None}
    j=int(np.argmax(loss))
    if j in (0,len(rows)-1) or not loss[j]>max(loss[j-1],loss[j+1]):
        return {'status':'FAIL','reason':'NO_STRICT_INTERIOR_PEAK','observed_interior_peak_nm':None}
    peak=float(wavelengths[j]);err=abs(peak-arrow_nm)/arrow_nm
    return {'status':'NEEDS_EVIDENCE','observed_interior_peak_nm':peak,
            'resonance_relative_error':err,'within_declared_5_percent':bool(err<=.05),
            'reason':'REQUIRES_SPECTRAL_REFINEMENT_AND_INDEPENDENT_MESH_PML_WINDOW'}
