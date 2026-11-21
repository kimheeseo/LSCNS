"""Lossless, indexed bundles for raw FE candidate fields; no numerical rounding."""
import argparse,hashlib,json,zipfile
from pathlib import Path
from modal_search_evidence import write_json


def package(root):
    root=Path(root);files=sorted(root.rglob('candidates.json'))
    files+=sorted((root/'acquisition_sources').glob('*.py'))
    if not files:raise ValueError('no actual candidate files to package')
    entries=[]
    with zipfile.ZipFile(root/'candidates.zip','w',compression=zipfile.ZIP_DEFLATED,compresslevel=9) as z:
        for p in files:
            raw=p.read_bytes();name=p.relative_to(root).as_posix()
            info=zipfile.ZipInfo(name,date_time=(2026,11,21,0,0,0));info.compress_type=zipfile.ZIP_DEFLATED
            z.writestr(info,raw,compress_type=zipfile.ZIP_DEFLATED,compresslevel=9)
            entries.append({'path':name,'sha256':hashlib.sha256(raw).hexdigest(),'bytes':len(raw)})
    index={'schema':'LSCNS_physics_field_bundle_v1','archive':'candidates.zip',
           'archive_sha256':hashlib.sha256((root/'candidates.zip').read_bytes()).hexdigest(),
           'entries':entries,'scope':'lossless actual candidate JSON; rejected and accepted fields retained'}
    write_json(root/'candidate_index.json',index);return index


def compact_summary(root):
    """Keep scalar mesh diagnostics in the review table; full arrays stay in the indexed bundle."""
    path=Path(root)/'summary.json'
    def compact(value):
        if isinstance(value,list):return [compact(x) for x in value]
        if not isinstance(value,dict):return value
        result={}
        for key,v in value.items():
            if key=='mesh_quality' and isinstance(v,dict):
                result[key]={k:n for k,n in v.items() if isinstance(n,(str,int,float,bool)) or n is None}
                result[key]['full_quality_arrays_in_candidate_bundle']=True
            else:result[key]=compact(v)
        return result
    write_json(path,compact(json.loads(path.read_text())))


def unpack(root):
    root=Path(root);index=json.loads((root/'candidate_index.json').read_text());archive=root/index['archive']
    if hashlib.sha256(archive.read_bytes()).hexdigest()!=index['archive_sha256']:raise ValueError('archive hash mismatch')
    with zipfile.ZipFile(archive) as z:
        if sorted(z.namelist())!=sorted(e['path'] for e in index['entries']):raise ValueError('archive member mismatch')
        checked=[]
        for e in index['entries']:
            p=(root/e['path']).resolve()
            if not p.is_relative_to(root.resolve()):raise ValueError('unsafe archive path')
            raw=z.read(e['path'])
            if len(raw)!=e['bytes'] or hashlib.sha256(raw).hexdigest()!=e['sha256']:raise ValueError('field member hash mismatch')
            checked.append((p,raw))
        for p,raw in checked:p.parent.mkdir(parents=True,exist_ok=True);p.write_bytes(raw)
    return len(checked)


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('action',choices=('package','unpack'));p.add_argument('root');a=p.parse_args()
    if a.action=='package':
        print(package(a.root))
        if (Path(a.root)/'summary.json').exists():compact_summary(a.root)
    else:print(unpack(a.root))
