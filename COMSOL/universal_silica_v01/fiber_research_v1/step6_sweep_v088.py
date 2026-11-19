"""STEP6e deterministic grid search. Pareto: all objectives no worse and >=1 strictly better.
Objectives carry explicit units and min/max direction; failed/nonfinite points never win.
No Bayesian/genetic optimizer or interpolated objective is silently substituted for FEM.
"""
import copy,itertools,json,math
from pathlib import Path
try:
 from .step4_v086 import digest
except ImportError:
 from step4_v086 import digest
VERSION='0.8.8'

def set_path(design,path,value):
 parts=path.split('.')
 if not parts or any(not p or p.startswith('_') for p in parts):raise ValueError('public existing dotted parameter path required')
 d=design
 for key in parts[:-1]:d=d[key]
 if parts[-1] not in d:raise ValueError('parameter path must already exist')
 if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value):raise ValueError('finite numeric sweep values required')
 d[parts[-1]]=value

def grid_sweep(design,axes,objectives,evaluate,constraints=None,max_cases=100,checkpoint=None,evaluator_id=None):
 if not axes or not objectives or not isinstance(evaluator_id,str) or not evaluator_id.strip():raise ValueError('axes, objectives and versioned evaluator_id required')
 for path,values in axes.items():
  if not isinstance(values,list) or not values or len(set(values))!=len(values):raise ValueError('nonempty unique axis values required')
  for v in values:set_path(copy.deepcopy(design),path,v)
 for name,spec in objectives.items():
  if spec.get('direction') not in ('min','max') or not spec.get('unit'):raise ValueError('objective direction and unit required')
 constraints=constraints or {}
 for name,spec in constraints.items():
  if not set(spec)<= {'min','max'} or not spec or any(not math.isfinite(v) for v in spec.values()):raise ValueError('finite constraint min/max required')
 count=math.prod(len(v) for v in axes.values())
 if not 1<=count<=max_cases<=1000:raise ValueError('grid exceeds explicit case budget (maximum 1000)')
 config={'design':design,'axes':axes,'objectives':objectives,'constraints':constraints,'evaluator_id':evaluator_id,'schema_version':'1.0.0','engine_version':VERSION};signature=digest(config)
 rows=[]
 if checkpoint and Path(checkpoint).exists():
  old=json.loads(Path(checkpoint).read_text())
  if old.get('config_sha256')!=signature:raise ValueError('checkpoint input/evaluator mismatch')
  rows=old['cases']
  for i,row in enumerate(rows):
   if row.get('case_index')!=i:raise ValueError('corrupt checkpoint order')
 root={'schema':'LSCNS_step6_grid_v088','schema_version':'1.0.0','engine_version':VERSION,'config_sha256':signature,'configuration':config,'cases':rows,'certified':False}
 keys=list(axes)
 for i,values in enumerate(itertools.product(*axes.values())):
  params=dict(zip(keys,values));d=copy.deepcopy(design)
  for path,v in params.items():set_path(d,path,v)
  if i<len(rows):
   if rows[i]['parameters']!=params or rows[i]['design_sha256']!=digest(d):raise ValueError('corrupt checkpoint parameters')
   continue
  row={'case_index':i,'parameters':params,'design_sha256':digest(d),'design':d}
  try:
   metrics=evaluate(d)
   for name in set(objectives)|set(constraints):
    if name not in metrics or isinstance(metrics[name],bool) or not isinstance(metrics[name],(int,float)) or not math.isfinite(metrics[name]):raise ValueError('missing/nonfinite objective or constraint: '+name)
   feasible=all(('min' not in s or metrics[k]>=s['min']) and ('max' not in s or metrics[k]<=s['max']) for k,s in constraints.items())
   row.update(status='COMPUTED',metrics=metrics,feasible=feasible)
  except Exception as err:
   row.update(status='FAILED',error=str(err),feasible=False)
  rows.append(row)
  if checkpoint:
   Path(checkpoint).parent.mkdir(parents=True,exist_ok=True);tmp=Path(str(checkpoint)+'.tmp');tmp.write_text(json.dumps(root,ensure_ascii=False,indent=2,allow_nan=False)+'\n');tmp.replace(checkpoint)
 feasible=[r for r in rows if r['status']=='COMPUTED' and r['feasible']]
 def score(r):return [r['metrics'][k]*(1 if s['direction']=='min' else -1) for k,s in objectives.items()]
 def dominates(a,b):return all(x<=y for x,y in zip(a,b)) and any(x<y for x,y in zip(a,b))
 front=[r['case_index'] for r in feasible if not any(dominates(score(q),score(r)) for q in feasible)]
 best={k:min(feasible,key=lambda r:r['metrics'][k]*(1 if s['direction']=='min' else -1))['case_index'] for k,s in objectives.items()} if feasible else {}
 root.update(pareto_case_indices=front,best_per_objective=best,best_designs={k:rows[i]['design'] for k,i in best.items()},
             limitation='finite grid only; objective accuracy inherits FEM validation; no guaranteed global optimum')
 if checkpoint:Path(checkpoint).write_text(json.dumps(root,ensure_ascii=False,indent=2,allow_nan=False)+'\n')
 return root
