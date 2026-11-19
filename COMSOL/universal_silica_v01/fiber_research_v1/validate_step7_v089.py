"""STEP7 V1..V9. PASS/FAIL/NEEDS_EVIDENCE are physics outcomes, not CI exit codes.
No item is complete until its stated checks pass. Missing evidence is never zero error.
A green software workflow does not override all_required_passed=False.
"""
import argparse,json,copy,hashlib,platform,importlib.metadata
from datetime import datetime,timezone
from pathlib import Path
import numpy as np
from validation_photonics_v089 import bv_curve,pcf_fsm_study,arf_resonance_study
from validation_reference_v089 import references
from validation_csv_v089 import parse_comsol_csv,compare_tables
from physical_metrics import confinement_loss,field_area_mfd
from step4_v086 import dispersion
from step3_reference import he11_exact
from validate_step3 import benchmark_design
from geometry import build_mesh
from step4_v086 import digest
VERSION='0.8.9'

def write(path,data):Path(path).write_text(json.dumps(data,ensure_ascii=False,indent=2,allow_nan=False)+'\n')
def item(key,title,status,scope,tolerances,evidence,needed=None):
 return dict(id=key,title=title,status=status,scope=scope,tolerances=tolerances,evidence=evidence,needed=needed,complete=status=='PASS')

def unit_evidence():
 checks=[]
 v=confinement_loss(1e-8,1550);expected=(20/np.log(10))*(2*np.pi/(1550e-9))*1e-8
 checks.append({'name':'loss_SI_dB_m_and_km','error_relative':abs(v['alpha_total_dB_m']/expected-1),'passed':abs(v['alpha_total_dB_m']/expected-1)<1e-12 and v['alpha_total_dB_km']==v['alpha_total_dB_m']*1000})
 checks.append({'name':'negative_Im_not_abs','passed':confinement_loss(-1e-8,1550)['alpha_total_dB_m']<0})
 lam=np.linspace(1500,1600,9);a=1e-8;n=1.45+a*(lam-1550)**2;r=dispersion(lam,n,np.ones(len(lam)-1));expected=-lam/299792458.*1e15*2*a
 err=float(np.max(abs(np.asarray(r['D_ps_nm_km'])-expected)));checks.append({'name':'D_negative_for_positive_n_curvature','maximum_absolute_error_ps_nm_km':err,'passed':bool(np.all(np.asarray(r['D_ps_nm_km'])<0) and err<1e-5)})
 x=np.linspace(-40,40,201);xx,yy=np.meshgrid(x,x);w=5.;e=np.exp(-(xx*xx+yy*yy)/w**2);ae=field_area_mfd(x,x,e,np.zeros_like(e))['Aeff_um2']
 err=abs(ae/(np.pi*w*w)-1);checks.append({'name':'Gaussian_Aeff_pi_w_squared','actual_um2':ae,'reference_um2':np.pi*w*w,'relative_error':float(err),'passed':err<1e-10})
 return checks

def csv_evidence():
 a='% Model: synthetic parser-only fixture\n% lambda (um),mode_id,real(ewfd.neff) (1),imag(ewfd.neff) (1)\n1.55,HE11_a,1.445,-3.0D-8\n'
 b='wavelength_nm,mode_id,neff\n1550,HE11_a,1.445+3.0e-8i\n'
 aa=parse_comsol_csv(a,-1);bb=parse_comsol_csv(b,1);r=compare_tables(aa,bb,1.,.01,0.)
 failures=[]
 for source in ('lambda (nm),neff\n1550,__import__("os")\n','wavelength_nm,mode_id,neff_real,neff_imag\n1550,a,NaN,1\n'):
  try:parse_comsol_csv(source,1);failures.append(False)
  except ValueError:failures.append(True)
 return {'status':'PASS' if r['status']=='PASS' and all(failures) else 'FAIL','fixture_comparison':r,'malformed_rejected':all(failures),'scope':'synthetic CSV parser/comparison regression only; not uploaded COMSOL experimental evidence'}

def sensitivity(step3):
 radiative=step3['V5']['cases'];v4=[];v5=[]
 for name in ('thickness_um','strength','order'):
  cases=[r for r in radiative if r['name'].startswith('pml_'+name+'_')];im=np.array([v for r in cases if r.get('state')=='COMPUTED' for v in r['neff_imag']]);complete=len(cases)==3 and all(r.get('state')=='COMPUTED' and r['tracking']['status']=='PASS_TRACKING' for r in cases)
  spread=float(np.ptp(im)/abs(np.mean(im))) if len(im) and np.mean(im)!=0 else None
  passed=complete and spread is not None and spread<=.05 and np.min(im)>1e-10
  v4.append({'parameter':name,'cases':cases,'Im_fractional_range':spread,'passed':bool(passed)})
 for label,source in (('guided',step3['V1']['cases']),('radiative',radiative)):
  rows=[copy.deepcopy(r) for r in source if r['name'].startswith(label+'_mesh_')]
  rows.sort(key=lambda r:int(r['name'].split('_')[-1]));last_re=None;last_im=None
  for row in rows:
   count=int(row['name'].split('_')[-1]);d=benchmark_design('guided' if label=='guided' else 'radiative');d['mesh'].update(elements_per_wavelength=count,curvature_segments={4:144,6:192,8:240}[count])
   if digest(d)!=row['input_digest']:raise ValueError('mesh quality reconstruction input hash mismatch')
   mesh=build_mesh(d);row['measured_mesh_quality']=mesh['quality'];row['mesh_quality_rebuild_triangle_count_matches']=mesh['quality']['triangles']==row['triangles']
  complete=len(rows)==3 and all(r.get('state')=='COMPUTED' and r['tracking']['status']=='PASS_TRACKING' and r['mesh_quality_rebuild_triangle_count_matches'] for r in rows)
  if complete:
   last_re=abs(float(np.mean(rows[-1]['neff_real'])-np.mean(rows[-2]['neff_real'])));den=abs(float(np.mean(rows[-1]['neff_imag'])))
   last_im=abs(float(np.mean(rows[-1]['neff_imag'])-np.mean(rows[-2]['neff_imag'])))/den if den>1e-10 else None
  passed=complete and last_re is not None and last_re<=1e-5 and (label=='guided' or last_im is not None and last_im<=.05)
  v5.append({'family':label,'cases':rows,'last_Re_change':last_re,'last_Im_relative_change':last_im,'guided_Im_status':'NOT_LEAKAGE_CERTIFIED_NEAR_ZERO' if label=='guided' else None,'passed':bool(passed)})
 return v4,v5

def build_report(step3_path,output,run_photonics=True):
 out=Path(output);out.mkdir(parents=True,exist_ok=True);ref=references();step3=json.loads(Path(step3_path).read_text());items=[]
 curve=bv_curve();write(out/'bv_curve.json',curve);v1pass=len(curve['cases'])>=4 and all(r['status']=='COMPUTED' and r['neff_error_ppm']<=10 for r in curve['cases'])
 items.append(item('V1','step-index b–V / exact vector HE11','PASS' if v1pass else 'FAIL','explicit synthetic circular HE11, finite V interval only',{'neff_error_ppm':10},curve))
 items.append(item('V2','COMSOL microstructured optical fiber','NEEDS_EVIDENCE','10 published PML mode values verified; geometry and numeric reproduction not established',{'Re_ppm':None,'Im_relative':None,'basis':'must be declared before reproduction, not inferred from reference rounding'},ref['V2'],ref['V2']['missing']))
 items.append(item('V3','COMSOL step-index fiber bend','NEEDS_EVIDENCE','current 2D conformal approximation differs from COMSOL axisymmetric eigenvalue normalization',{'complex_neff':None,'loss_dB_m':None},ref['V3'],ref['V3']['missing']))
 pml,meshes=sensitivity(step3)
 items.append(item('V4','PML independence','PASS' if all(r['passed'] for r in pml) else 'FAIL','synthetic strongly radiating circular interface; not tiny-loss ARF certification',{'Im_fractional_range':.05,'minimum_positive_Im':1e-10,'basis':'declared engineering acceptance'},pml))
 items.append(item('V5','mesh convergence','PASS' if all(r['passed'] for r in meshes) else 'FAIL','same physical modes tracked by actual FE subspace; actual edge size and triangles recorded',{'last_absolute_Re_change':1e-5,'radiative_last_relative_Im_change':.05},meshes))
 fsm=pcf_fsm_study() if run_photonics else {'status':'NEEDS_EVIDENCE','cases':[],'missing':'actual FSM/FM study not run'};write(out/'pcf_fsm.json',fsm)
 items.append(item('V6','solid PCF / FSM-based ESM','NEEDS_EVIDENCE','computed FSM candidates and primary-paper V_PCF=pi criterion; no ESM proof',{'FSM_neff_convergence':None,'reference_V_error':None,'basis':'no fabricated primary numerical tolerance'},fsm,fsm['missing']))
 arf=arf_resonance_study() if run_photonics else {'status':'NEEDS_EVIDENCE','cases':[],'missing':'actual tracked ARF scan not run'};write(out/'arf_resonance.json',arf)
 items.append(item('V7','single-ring ARF / ARROW resonance',arf['status'],'actual vector scan and analytic reference shown separately; no formula-only PASS',{'peak_relative_wavelength':.05,'mesh_PML_convergence':'required'},arf,arf['missing']))
 unit=unit_evidence();items.append(item('V8','units / formulas','PASS' if all(r['passed'] for r in unit) else 'FAIL','independent analytic identities, not FEM accuracy',{'loss_relative':1e-12,'D_absolute_ps_nm_km':1e-5,'Gaussian_Aeff_relative':1e-10},unit))
 csvr=csv_evidence();items.append(item('V9','COMSOL CSV import / error table',csvr['status'],'synthetic parser/comparison regression only; upload actual COMSOL exports for external comparison',{'fixture_Re_ppm':1,'fixture_Im_relative':.01,'fixture_Im_absolute':0},csvr))
 files=['vector_fem.py','vector_solver_v085.py','vector_spectrum.py','geometry.py','materials.py','physical_metrics.py','validation_photonics_v089.py','validate_step7_v089.py','validation_csv_v089.py']
 hashes={name:hashlib.sha256((Path(__file__).parent/name).read_bytes()).hexdigest() for name in files}
 report={'schema':'LSCNS_step7_validation_v089','schema_version':'1.0.0','engine_version':VERSION,'modal_solver_version':'0.8.6','timestamp_utc':datetime.now(timezone.utc).isoformat(),'source_file_sha256':hashes,'source_step3_report_sha256':hashlib.sha256(Path(step3_path).read_bytes()).hexdigest(),'environment':{'python':platform.python_version(),**{k:importlib.metadata.version(k) for k in ('numpy','scipy','scikit-fem','gmsh')}},'references':ref,'items':items,'all_required_passed':all(r['status']=='PASS' for r in items),'completion_label':'NOT_COMPLETE_UNTIL_ALL_V1_V9_PASS','step8_started':False,'next_proposal':'obtain V2 original parameters and V3 normalized bend CSV; refine/converge FSM and ARF modal spectra before physical validation claims'}
 write(out/'validation.json',report);return report

def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('--step3-report',required=True);p.add_argument('--out',required=True);p.add_argument('--skip-photonics',action='store_true');a=p.parse_args();r=build_report(a.step3_report,a.out,not a.skip_photonics);print({v['id']:v['status'] for v in r['items']})
 if any(v['status']!='PASS' for v in r['items'] if v['id'] in ('V1','V4','V5','V8','V9')):raise SystemExit(1)
if __name__=='__main__':main()
