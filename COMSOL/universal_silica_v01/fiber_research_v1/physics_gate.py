"""Fail-closed evaluation of physical evidence, separate from software exit codes.

Evidence file hashes bind declared records to artifacts; hashes do not prove the
truth of external provenance or a scientific review. No missing value becomes 0.
"""
import argparse
import csv
import copy
import hashlib
import io
import json
import math
import zipfile
from pathlib import Path
from datetime import datetime,timezone
from validation_csv_v089 import parse_comsol_csv,compare_tables
from modal_search_evidence import peak_evidence,write_json
from bend_reference import normalize_bend
from step4_v086 import digest
from step6_fields_v088 import subspace_overlap


def normalized_bent_csv(raw,normalization):
    """Canonical export contract; values and evaluated radii remain full precision."""
    out=io.StringIO();writer=csv.writer(out)
    writer.writerow(['wavelength_nm','mode_id','neff_real','neff_imag'])
    rows=list(csv.DictReader(io.StringIO(raw)))
    if not rows:raise ValueError('empty bent export')
    basis=normalization['comparison_radius_basis']
    for row in rows:
        result=normalize_bend(complex(float(row['raw_real']),float(row['raw_imag'])),
                representation=normalization['representation'],wavelength_m=float(row['wavelength_nm'])*1e-9,
                r_average_m=float(row['r_average_m']),r_power_m=float(row['r_power_m']),
                imag_sign=normalization['imag_sign'])
        writer.writerow([row['wavelength_nm'],row['mode_id'],result['neff_'+basis+'_real'],result['neff_'+basis+'_imag']])
    return out.getvalue()


def artifact(root,spec):
    if not isinstance(spec,dict) or not spec.get('path') or not spec.get('sha256'):
        raise ValueError('missing artifact path/SHA256')
    root=Path(root).resolve();p=(root/spec['path']).resolve()
    if not p.is_relative_to(root):raise ValueError('artifact outside evidence root')
    raw=p.read_bytes()
    if hashlib.sha256(raw).hexdigest()!=spec['sha256']:raise ValueError('artifact SHA256 mismatch: '+spec['path'])
    return raw.decode('utf-8')


def tolerance(value):
    if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value) or value<0:
        raise ValueError('finite predeclared nonnegative tolerance required')
    return value


def identity(root, spec):
    r=json.loads(artifact(root,spec));pairs=r.get('mode_mapping',[])
    if not pairs or not r.get('reviewed_by') or not r.get('basis'):
        raise ValueError('reviewed field/subspace identity mapping required')
    for pair in pairs:
        # Every mapping must bind actual reference and calculated fields, not just neff sorting.
        fields=[]
        for key in ('reference_field','calculated_field'):
            raw=json.loads(artifact(root,pair.get(key)))
            modes=raw['modes'] if 'modes' in raw else [raw]
            fields.append(modes)
        corr=pair.get('minimum_subspace_correlation')
        if not isinstance(corr,(int,float)) or not math.isfinite(corr) or not .8<=corr<=1:
            raise ValueError('unqualified modal identity')
        measured=subspace_overlap(fields[0],fields[1])
        if not measured or min(measured)<.8:raise ValueError('actual field/subspace mapping rejected')
    return r


def independent_convergence(root,spec):
    r=json.loads(artifact(root,spec));axes=r.get('axes',{})
    limits=r.get('tolerances',{})
    re_tol=tolerance(limits.get('Re_absolute'))
    im_rel=tolerance(limits.get('Im_relative'));im_abs=tolerance(limits.get('Im_absolute'))
    for axis in ('mesh','window','pml_thickness','pml_strength','pml_order'):
        rows=axes.get(axis,[])
        if len(rows)<3:raise ValueError('three independent levels required: '+axis)
        if len({str(v.get('level')) for v in rows})<3:raise ValueError('distinct levels required: '+axis)
        evaluated=[];fixed=None;previous=None
        for row in rows:
            raw=json.loads(artifact(root,row.get('result')));design=raw['input']
            if row.get('input_sha256')!=digest(design):raise ValueError('convergence input hash mismatch')
            pool=raw.get('candidates',raw.get('modes',[]));indices=row.get('mode_indices',[])
            if not indices or len(set(indices))!=len(indices):raise ValueError('explicit independent mode indices required')
            modes=[pool[i] for i in indices]
            if any(m.get('status')!='candidate_unverified' for m in modes):raise ValueError('unqualified convergence mode')
            corr=subspace_overlap(previous,modes) if previous else [1.]*len(modes)
            if len(corr)!=len(modes) or min(corr)<.8:raise ValueError('convergence branch identity missing')
            stable=copy.deepcopy(design)
            if axis=='mesh':stable.pop('mesh',None)
            elif axis=='window':stable.pop('domain_radius_um',None)
            else:stable['boundary'].pop({'pml_thickness':'thickness_um','pml_strength':'strength','pml_order':'order'}[axis],None)
            if fixed is not None and stable!=fixed:raise ValueError('coupled changes in independent sweep: '+axis)
            fixed=stable;previous=modes
            evaluated.append({'neff_real':sum(m['neff_real'] for m in modes)/len(modes),
                              'neff_imag':sum(m['neff_imag'] for m in modes)/len(modes)})
        if len({row['input_sha256'] for row in rows})!=len(rows):raise ValueError('distinct actual input levels required')
        for a,b in zip(evaluated[-3:-1],evaluated[-2:]):
            ar,ai,br,bi=(v for row in (a,b) for v in (row.get('neff_real'),row.get('neff_imag')))
            if not all(isinstance(v,(int,float)) and math.isfinite(v) for v in (ar,ai,br,bi)):
                raise ValueError('finite convergence eigenvalues required')
            if ai<=0 or bi<=0:raise ValueError('positive leakage reference required; near zero needs explicit separate policy')
            if abs(br-ar)>re_tol or abs(bi-ai)>max(im_abs,im_rel*abs(bi)):
                return False
    return True


def external_gate(key,manifest,root):
    try:
        required=['reference_csv','calculated_csv','identity','convergence','model_manifest']
        if key=='V2':required+=['exact_parameters']
        else:required+=['raw_bent_csv','normalization_manifest','formulation_review']
        for name in required:artifact(root,manifest.get(name))
        model=json.loads(artifact(root,manifest['model_manifest']))
        if not model.get('comsol_version') or not model.get('source_url') or not model.get('input_equivalence_reviewed_by'):
            raise ValueError('version/source/input equivalence review required')
        mapping=identity(root,manifest['identity'])
        if not independent_convergence(root,manifest['convergence']):
            return {'status':'FAIL','reason':'INDEPENDENT_CONVERGENCE_FAILED'}
        if key=='V3':
            norm=json.loads(artifact(root,manifest['normalization_manifest']))
            for name in ('representation','length_unit','imag_sign','comparison_radius_basis'):
                if norm.get(name) is None:raise ValueError('bend normalization missing: '+name)
            if norm['length_unit']!='m' or norm['comparison_radius_basis'] not in ('geometry','power'):
                raise ValueError('explicit SI normalization required')
            review=json.loads(artifact(root,manifest['formulation_review']))
            if not review.get('reviewed_by') or review.get('equivalent_comparison') is not True:
                raise ValueError('axisymmetric/conformal formulation equivalence not established')
        t=manifest.get('tolerances',{})
        ref=parse_comsol_csv(artifact(root,manifest['reference_csv']),manifest.get('reference_imag_sign'))
        calc=parse_comsol_csv(artifact(root,manifest['calculated_csv']),manifest.get('calculated_imag_sign'))
        if key=='V3':
            normalized=parse_comsol_csv(normalized_bent_csv(artifact(root,manifest['raw_bent_csv']),norm),1)
            consistent=compare_tables(normalized,ref,1e-6,1e-10,1e-18)
            if consistent['status']!='PASS':raise ValueError('raw bent export and normalized reference disagree')
        mapped={(p.get('wavelength_nm'),p.get('reference_mode_id'),p.get('calculated_mode_id')) for p in mapping['mode_mapping']}
        if any((r['wavelength_nm'],r['mode_id'],r['mode_id']) not in mapped for r in ref['rows']):
            raise ValueError('explicit shared mode labels and full wavelength mapping required')
        if key=='V2' and len(ref['rows'])!=10:raise ValueError('all ten published V2 modes required')
        comparison=compare_tables(calc,ref,tolerance(t.get('Re_ppm')),tolerance(t.get('Im_relative')),tolerance(t.get('Im_absolute')))
        return {'status':comparison['status'],'comparison':comparison,
                'scope':'input-matched reviewed external benchmark only; not manufacturer certification'}
    except (ValueError,TypeError,KeyError,OSError) as exc:
        return {'status':'NEEDS_EVIDENCE','reason':str(exc),
                'missing_artifacts':[name for name in required if not manifest.get(name)]}


def evaluate(root,baseline):
    root=Path(root)
    if sorted(v['id'] for v in baseline['items'])!=['V'+str(i) for i in range(1,10)]:
        raise ValueError('exactly V1..V9 baseline items required')
    items={v['id']:{'status':v['status'],'scope':v['scope'],
            'basis':'retained historical scoped benchmark, not recomputed in this evaluation'} for v in baseline['items']}
    manifest=json.loads((root/'external_manifest.json').read_text())
    actual=hashlib.sha256(json.dumps(baseline,sort_keys=True,separators=(',',':'),ensure_ascii=False).encode()).hexdigest()
    if actual!=manifest.get('baseline_canonical_sha256'):raise ValueError('baseline evidence hash mismatch')
    bundles=[]
    for p in sorted(root.rglob('candidate_index.json')):
        index=json.loads(p.read_text());archive=p.parent/index['archive']
        if hashlib.sha256(archive.read_bytes()).hexdigest()!=index['archive_sha256']:
            raise ValueError('candidate archive SHA256 mismatch')
        with zipfile.ZipFile(archive) as z:
            if sorted(z.namelist())!=sorted(v['path'] for v in index['entries']):raise ValueError('candidate member mismatch')
            for member in index['entries']:
                raw=z.read(member['path'])
                if len(raw)!=member['bytes'] or hashlib.sha256(raw).hexdigest()!=member['sha256']:
                    raise ValueError('candidate field SHA256 mismatch')
        bundles.append({'archive':str(archive.relative_to(root)),'verified_members':len(index['entries']),
                        'sha256':index['archive_sha256']})
    for key in ('V2','V3'):items[key]=external_gate(key,manifest.get(key,{}),root)
    p6=root/'V6/summary.json'
    if p6.exists():
        r=json.loads(p6.read_text());changes=[]
        for case in r['cases']:
            fsm=[v['n_FSM'] for v in case['FSM']]
            changes.append(abs(fsm[-1]-fsm[-2]) if len(fsm)>=2 else None)
        items['V6']={'status':'NEEDS_EVIDENCE','FSM_last_changes':changes,
                     'reason':'HOM cutoff/primary quantitative reference and independent full-section convergence remain unqualified',
                     'source':'V6/summary.json'}
        try:
            q=manifest.get('V6',{});identity(root,q.get('identity'))
            if not independent_convergence(root,q.get('convergence')):
                items['V6'].update(status='FAIL',reason='FULL_SECTION_CONVERGENCE_FAILED')
            else:
                qualification=json.loads(artifact(root,q.get('HOM_reference_comparison')))
                reference_data=json.loads(artifact(root,qualification.get('primary_reference_data')))
                artifact(root,qualification.get('HOM_field_evidence'))
                if not qualification.get('reviewed_by') or not qualification.get('cutoff_bracket'):
                    raise ValueError('reviewed HOM cutoff bracket required')
                limit=tolerance(qualification.get('reference_V_relative_tolerance'))
                comparisons=[]
                for case in r['cases']:
                    matches=[v for v in reference_data.get('rows',[]) if
                             v.get('wavelength_nm')==case['wavelength_nm'] and
                             v.get('pitch_um')==case['pitch_um'] and v.get('diameter_ratio')==case['diameter_ratio']]
                    if len(matches)!=1:raise ValueError('unique input-matched primary V reference required')
                    reference=matches[0];value=reference.get('V_PCF')
                    uncertainty=tolerance(reference.get('uncertainty_absolute'))
                    if not isinstance(value,(int,float)) or not math.isfinite(value) or value<=0:
                        raise ValueError('positive finite primary V reference required')
                    fm=case['FM_mesh'][-1]['neff_real'];fm=sum(fm)/len(fm);nf=case['FSM'][-1]['n_FSM']
                    if fm<=nf:raise ValueError('FM must exceed converged FSM for V_PCF')
                    calculated=2*math.pi*case['pitch_um']/(case['wavelength_nm']*.001)*math.sqrt(fm*fm-nf*nf)
                    comparisons.append({'wavelength_nm':case['wavelength_nm'],'V_calculated':calculated,
                                        'V_reference':value,'reference_uncertainty_absolute':uncertainty,
                                        'relative_error_beyond_uncertainty':max(0.,abs(calculated-value)-uncertainty)/value})
                error=max(v['relative_error_beyond_uncertainty'] for v in comparisons)
                items['V6']['reference_comparison']=comparisons
                n_tol=tolerance(r['policy']['FSM_absolute_neff_last_two_changes'])
                f_tol=tolerance(r['policy']['FM_absolute_neff_last_two_changes'])
                from run_physics_v6 import convergence
                stable=True
                for case in r['cases']:
                    stable &= convergence([v['n_FSM'] for v in case['FSM']],n_tol)
                    mesh=case['FM_mesh']
                    stable &= len(mesh)>=3 and all(v['status']=='TRACKED_CANDIDATE' for v in mesh)
                    if stable:stable &= convergence([sum(v['neff_real'])/len(v['neff_real']) for v in mesh],f_tol)
                items['V6'].update(status='PASS' if stable and error<=limit else 'FAIL',
                    reason='FINITE_PCF_BENCHMARK_EVALUATED_NOT_ESM_PROOF')
        except (ValueError,TypeError,KeyError,OSError) as exc:items['V6']['reason']=str(exc)
    p7=root/'V7_adaptive/summary.json'
    if p7.exists():
        r=json.loads(p7.read_text());peak=peak_evidence(r['cases'],r['arrow_lambda_nm'])
        items['V7']={'status':peak['status'],'peak':peak,'source':'V7_adaptive/summary.json'}
        if peak['status']=='NEEDS_EVIDENCE':
            # A detected peak is only an intermediate outcome. Independent records must be supplied.
            try:
                q=manifest.get('V7',{});identity(root,q.get('identity'))
                converged=independent_convergence(root,q.get('convergence'))
                refinement=json.loads(artifact(root,q.get('spectral_refinement')))
                peaks=[];local_steps=[]
                levels=refinement.get('levels',[])
                if len(levels)<3:raise ValueError('three raw spectral refinement levels required')
                for level in levels:
                    spectrum=json.loads(artifact(root,level.get('spectrum')))
                    observed=peak_evidence(spectrum['cases'],r['arrow_lambda_nm'])
                    if observed['status']=='FAIL':
                        items['V7']['status']='FAIL'
                        raise ValueError('refined raw spectrum has a broken branch or no interior peak')
                    peaks.append(observed['observed_interior_peak_nm'])
                    rows=spectrum['cases'];j=max(range(len(rows)),key=lambda i:rows[i]['alpha_total_dB_m'])
                    local_steps.append(max(rows[j]['wavelength_nm']-rows[j-1]['wavelength_nm'],
                                           rows[j+1]['wavelength_nm']-rows[j]['wavelength_nm']))
                if any(b>=a for a,b in zip(local_steps,local_steps[1:])):
                    raise ValueError('spectral grids must refine around the same interior peak')
                delta=tolerance(refinement.get('relative_tolerance'))
                if len(peaks)<3 or any(not isinstance(x,(int,float)) or not math.isfinite(x) or x<=0 for x in peaks):
                    raise ValueError('three positive refined spectral peaks required')
                stable=all(abs(b-a)/b<=delta for a,b in zip(peaks[-3:-1],peaks[-2:]))
                within=abs(peaks[-1]-r['arrow_lambda_nm'])/r['arrow_lambda_nm']<=.05
                items['V7']['status']='PASS' if converged and stable and within else 'FAIL'
                items['V7']['refined_peak_wavelength_nm']=peaks
            except (ValueError,TypeError,KeyError,OSError) as exc:items['V7']['reason']=str(exc)
    return {'schema':'LSCNS_physics_gate_v1','timestamp_utc':datetime.now(timezone.utc).isoformat(),
            'items':items,'all_required_passed':all(v['status']=='PASS' for v in items.values()),
            'software_success_is_physical_completion':False,'verified_bundles':bundles,
            'artifact_sha256':{str(p.relative_to(root)):hashlib.sha256(p.read_bytes()).hexdigest()
                               for p in sorted(root.rglob('*')) if p.is_file() and p.suffix in ('.json','.zip')
                               and p.name not in ('candidates.json','progress.json') and '_cache_sources' not in p.parts},
            'scope':'V1/V4/V5 historical synthetic scope; V8/V9 formula/parser scope; no product certification'}


if __name__=='__main__':
    a=argparse.ArgumentParser();a.add_argument('--evidence-root',required=True);a.add_argument('--baseline',required=True)
    a.add_argument('--out',required=True);a.add_argument('--require-all',action='store_true');args=a.parse_args()
    r=evaluate(args.evidence_root,json.loads(Path(args.baseline).read_text()));write_json(args.out,r)
    print({k:v['status'] for k,v in r['items'].items()})
    if args.require_all and not r['all_required_passed']:raise SystemExit(1)
