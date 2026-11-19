"""Result provenance contract. SHA256 hashes EXACT included canonical UTF-8 bytes.
Browser verifies canonical_input_json then structurally compares parsed input;
no incompatible reserialization of 1.0 / scientific notation / Unicode is needed.
"""
import copy,hashlib,json,platform
from datetime import datetime,timezone
from pathlib import Path
VERSION='0.8.10'
def canonical(value):return json.dumps(value,sort_keys=True,separators=(',',':'),ensure_ascii=False,allow_nan=False)
def stamp(result,source_input=None,solver_version=None):
 d=copy.deepcopy(source_input if source_input is not None else result.get('input',result.get('source_design')))
 if d is None:raise ValueError('source input required to stamp a result')
 text=canonical(d);result['input_sha256']=hashlib.sha256(text.encode()).hexdigest();result['solver_version']=solver_version or result.get('engine_version','unknown');result['app_version']=VERSION;result['timestamp_utc']=datetime.now(timezone.utc).isoformat();result['result_contract_version']='1.1.0'
 result['provenance']={'hash_algorithm':'SHA256','canonical_input_json':text,'encoding':'UTF-8 sorted keys compact JSON ensure_ascii=False','python_version':platform.python_version(),'physical_validation_upgraded':False,'hash_scope':'input_manifest' if result.get('input_manifest')==d else 'input'}
 return result
def verify(result):
 p=result.get('provenance',{});text=p.get('canonical_input_json')
 if not isinstance(text,str) or hashlib.sha256(text.encode()).hexdigest()!=result.get('input_sha256'):raise ValueError('input hash mismatch or missing provenance')
 d=result.get('input_manifest') if p.get('hash_scope')=='input_manifest' else result.get('input',result.get('source_design'))
 if p.get('hash_scope')=='input_manifest' and d.get('design')!=result.get('input'):raise ValueError('input manifest design differs from result input')
 if d!=json.loads(text):raise ValueError('canonical input differs from result input')
 if not result.get('solver_version') or not result.get('timestamp_utc'):raise ValueError('solver version/time required')
 try:
  dt=datetime.fromisoformat(result['timestamp_utc'].replace('Z','+00:00'))
  if dt.utcoffset() is None or dt.utcoffset().total_seconds()!=0:raise ValueError('UTC required')
 except (TypeError,ValueError,AttributeError):raise ValueError('valid UTC timestamp required') from None
 return True
