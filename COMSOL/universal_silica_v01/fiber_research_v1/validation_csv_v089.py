"""STEP7 V9: strict COMSOL CSV import and explicit-key comparison, no eval.
ppm=1e6*abs(calculated-reference)/abs(reference).
Im pass iff abs(delta)<=absolute_tolerance+relative_tolerance*abs(reference).
A positive imaginary part is passive for exp(i beta z-i omega t). COMSOL's
opposite time convention must be explicitly selected, never abs(Im).
Numeric table agreement does not establish matching physics/geometry or mode identity.
"""
import csv,io,re,math,json,argparse
from pathlib import Path
VERSION='0.8.9'
_FLOAT=r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eEdD][+-]?\d+)?'
def number(s):
 s=str(s).strip().replace('−','-')
 if not re.fullmatch(_FLOAT,s):raise ValueError('finite numeric cell required: '+s[:80])
 x=float(s.replace('D','e').replace('d','e'))
 if not math.isfinite(x):raise ValueError('nonfinite number')
 return x

def complex_cell(s):
 s=re.sub(r'\s+','',str(s).strip().replace('−','-'))
 if s.startswith('(') or s.endswith(')'):
  if not (s.startswith('(') and s.endswith(')')):raise ValueError('unbalanced complex literal parentheses')
  s=s[1:-1]
 if not s.endswith(('i','j')):return number(s),0.
 signed=r'[+-](?:\d+(?:\.\d*)?|\.\d+)(?:[eEdD][+-]?\d+)?'
 a=re.fullmatch('('+_FLOAT+')('+signed+')[ij]',s)
 if a:return number(a[1]),number(a[2])
 if re.fullmatch(_FLOAT+'[ij]',s):return 0.,number(s[:-1])
 raise ValueError('unsupported complex literal; expected a+bi')

def header_kind(h):
 h=h.strip().lstrip('%#').strip();s=re.sub(r'[^a-z0-9]','',h.lower())
 if s in ('mode','modeid','modenumber','modeindex','modeid1','mode1'):return 'mode_id',None
 if 'lambda' in s or 'wavelength' in s or 'lda0' in s:
  u='nm' if 'nm' in s else 'um' if 'um' in s or 'µm' in h else 'm' if re.search(r'\(m\)|\[m\]',h) else None
  return 'wavelength',u
 if ('imag' in s or s.startswith('im')) and ('neff' in s or 'eff' in s):return 'neff_imag',None
 if ('real' in s or s.startswith('re')) and ('neff' in s or 'eff' in s):return 'neff_real',None
 if s in ('neff','neff1','ewfdneff','ewfdneff1','effectiveindex'):return 'complex',None
 return None,None

def parse_comsol_csv(text,imag_sign,default_wavelength_nm=None,wavelength_unit=None,confirm_row_mode_ids=False):
 if imag_sign not in (-1,1) or isinstance(imag_sign,bool):raise ValueError('explicit imag_sign = +1 (our convention) or -1 (opposite convention) required')
 if len(text)>5_000_000:raise ValueError('CSV exceeds 5 MB')
 lines=text.lstrip('\ufeff').splitlines();header=None;delimiter=None;start=None
 for i,line in enumerate(lines):
  candidate=line.lstrip().lstrip('%#').strip()
  for sep in (',',';','\t'):
   row=next(csv.reader([candidate],delimiter=sep));kinds=[header_kind(h)[0] for h in row]
   if ('complex' in kinds or 'neff_real' in kinds):
    header=row;delimiter=sep;start=i+1;break
  if header is not None:break
 if header is None:raise ValueError('CSV header with neff and wavelength/mode columns not found')
 mapping={};unit=None
 for j,h in enumerate(header):
  kind,u=header_kind(h)
  if kind:
   if kind in mapping:raise ValueError('duplicate recognized column: '+kind)
   mapping[kind]=j
   if kind=='wavelength':unit=u or wavelength_unit
 if 'complex' not in mapping and not {'neff_real','neff_imag'}<=mapping.keys():raise ValueError('separate real column requires explicit imaginary column; no guessed zero loss')
 if 'complex' in mapping and ('neff_real' in mapping or 'neff_imag' in mapping):raise ValueError('ambiguous complex and separate neff columns')
 if 'wavelength' in mapping and unit not in ('nm','um','m'):raise ValueError('wavelength unit required (nm/um/m)')
 if 'wavelength' not in mapping and default_wavelength_nm is None:raise ValueError('single-wavelength CSV requires explicit default_wavelength_nm')
 if 'mode_id' not in mapping and not confirm_row_mode_ids:raise ValueError('mode IDs missing; confirm row-order mapping explicitly or add mode_id column')
 rows=[];counters={};seen=set()
 for line_no,line in enumerate(lines[start:],start+1):
  if not line.strip() or line.lstrip().startswith(('%','#')):continue
  values=next(csv.reader([line],delimiter=delimiter))
  if len(values)!=len(header):raise ValueError('row '+str(line_no)+': column count mismatch')
  lam=number(values[mapping['wavelength']])*{'nm':1.,'um':1000.,'m':1e9}[unit] if 'wavelength' in mapping else number(default_wavelength_nm)
  if lam<=0:raise ValueError('positive wavelength required')
  counters[lam]=counters.get(lam,0)+1
  mode=values[mapping['mode_id']].strip() if 'mode_id' in mapping else str(counters[lam])
  if not mode:raise ValueError('empty mode ID')
  key=(round(lam,7),mode)
  if key in seen:raise ValueError('duplicate wavelength/mode key')
  seen.add(key)
  reff,im=complex_cell(values[mapping['complex']]) if 'complex' in mapping else (number(values[mapping['neff_real']]),number(values[mapping['neff_imag']]))
  rows.append({'wavelength_nm':lam,'mode_id':mode,'neff_real':reff,'neff_imag':imag_sign*im,'raw_neff_imag':im,'source_line':line_no})
 if not rows:raise ValueError('CSV has no data rows')
 return {'schema':'LSCNS_comsol_csv_v089','schema_version':'1.0.0','engine_version':VERSION,'rows':rows,'imag_sign_applied':imag_sign,
  'mode_mapping':'EXPLICIT_MODE_IDS' if 'mode_id' in mapping else 'USER_CONFIRMED_ROW_ORDER_ONLY_NOT_PHYSICAL_ID',
  'scope':'numeric input only; units and phase convention explicitly supplied'}

def vector_to_table(result,mode_ids=None):
 if result.get('schema')!='LSCNS_fiber_research_vector_v1' or not result.get('modes'):raise ValueError('actual vector result JSON required')
 lam=number(result['input']['wavelength_nm'])
 if lam<=0:raise ValueError('positive wavelength required')
 ids=mode_ids or [str(i+1) for i in range(len(result['modes']))]
 if len(ids)!=len(result['modes']) or len(set(ids))!=len(ids) or any(not isinstance(v,str) or not v for v in ids):raise ValueError('unique mode IDs matching result count required')
 return {'schema':'LSCNS_comsol_csv_v089','schema_version':'1.0.0','engine_version':VERSION,
         'rows':[{'wavelength_nm':lam,'mode_id':mid,'neff_real':number(m['neff_real']),'neff_imag':number(m['neff_imag'])} for mid,m in zip(ids,result['modes'])],
         'mode_mapping':'USER_EXPLICIT_IDS' if mode_ids else 'ROW_INDEX_ONLY_NOT_PHYSICAL_ID',
         'source_solver_version':result.get('engine_version'),'source_input':result['input'],
         'physical_cross_validation_certified':False}

def compare_tables(calculated,reference,real_ppm_tolerance,imag_relative_tolerance,imag_absolute_tolerance,wavelength_tolerance_nm=1e-5):
 tolerances=[real_ppm_tolerance,imag_relative_tolerance,imag_absolute_tolerance,wavelength_tolerance_nm]
 if any(isinstance(v,bool) or not isinstance(v,(float,int)) or not math.isfinite(v) or v<0 for v in tolerances):raise ValueError('explicit finite nonnegative tolerances required')
 a=calculated['rows'];b=reference['rows'];rows=[];used=set()
 for ref in b:
  candidates=[(j,c) for j,c in enumerate(a) if c['mode_id']==ref['mode_id'] and abs(c['wavelength_nm']-ref['wavelength_nm'])<=wavelength_tolerance_nm]
  row={'reference':ref,'calculated':None,'status':'NEEDS_EVIDENCE','reason':'MISSING_OR_AMBIGUOUS_MATCH'}
  if len(candidates)==1:
   j,c=candidates[0]
   if j in used:raise ValueError('ambiguous wavelength matching; one calculated row reused')
   used.add(j);dr=abs(c['neff_real']-ref['neff_real']);di=abs(c['neff_imag']-ref['neff_imag']);ppm=dr/abs(ref['neff_real'])*1e6 if ref['neff_real'] else None
   limit=imag_absolute_tolerance+imag_relative_tolerance*abs(ref['neff_imag'])
   for v in (dr,di,limit):
    if not math.isfinite(v):raise ValueError('nonfinite comparison')
   valid=ppm is not None and ppm<=real_ppm_tolerance and di<=limit
   row.update(calculated=c,absolute_Re_error=dr,Re_error_ppm=ppm,absolute_Im_error=di,Im_error_relative=di/abs(ref['neff_imag']) if ref['neff_imag'] else None,
              Im_allowed_absolute_error=limit,status='PASS' if valid else 'FAIL',reason='EXPLICIT_KEYS_NUMERIC_COMPARISON')
  rows.append(row)
 unmatched=[c for j,c in enumerate(a) if j not in used]
 overall='FAIL' if any(r['status']=='FAIL' for r in rows) else 'NEEDS_EVIDENCE' if unmatched or any(r['status']!='PASS' for r in rows) or not rows else 'PASS'
 return {'schema':'LSCNS_validation_csv_comparison_v089','schema_version':'1.0.0','engine_version':VERSION,'status':overall,'rows':rows,'unmatched_calculated':unmatched,
 'tolerances':dict(zip(('Re_ppm','Im_relative','Im_absolute','wavelength_nm'),tolerances)),
 'physical_cross_validation_certified':False,'scope':'table agreement only; geometry/materials/boundaries/mode identity require separate evidence'}

def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('calculated');p.add_argument('reference');p.add_argument('--calculated-imag-sign',type=int,choices=[-1,1],required=True);p.add_argument('--reference-imag-sign',type=int,choices=[-1,1],required=True);p.add_argument('--wavelength-nm',type=float);p.add_argument('--unit',choices=['nm','um','m']);p.add_argument('--confirm-row-mode-ids',action='store_true');p.add_argument('--re-ppm',type=float,required=True);p.add_argument('--im-relative',type=float,required=True);p.add_argument('--im-absolute',type=float,required=True);p.add_argument('--out',required=True);v=p.parse_args()
 def load(path,sign):
  if Path(path).suffix.lower()=='.json':
   if not v.confirm_row_mode_ids:raise ValueError('vector JSON default mode indices require --confirm-row-mode-ids')
   if sign!=1:raise ValueError('our vector JSON convention already has passive positive Im')
   return vector_to_table(json.loads(Path(path).read_text()))
  return parse_comsol_csv(Path(path).read_text(),sign,v.wavelength_nm,v.unit,v.confirm_row_mode_ids)
 result=compare_tables(load(v.calculated,v.calculated_imag_sign),load(v.reference,v.reference_imag_sign),v.re_ppm,v.im_relative,v.im_absolute)
 Path(v.out).write_text(json.dumps(result,indent=2,allow_nan=False)+'\n');print(result['status'])
if __name__=='__main__':main()
