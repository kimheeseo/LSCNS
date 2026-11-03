from pathlib import Path
import sys,json,time,threading,urllib.request,urllib.error
import numpy as np
import pytest
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from engine import DEFAULTS,PRESETS,normalize,run_job
from launch import WorkbenchServer


@pytest.mark.parametrize('key,value',[('wavelength_nm',float('nan')),('core_elements',80.5),('mesh_size_um',0),('radius_um',-1),('model','unknown'),('delta_n',True)])
def test_invalid_values(key,value):
    with pytest.raises(ValueError):normalize({key:value})


def test_overlap_and_unsupported_v_rejected():
    with pytest.raises(ValueError):normalize({'model':'mcf','pitch_um':8.})
    with pytest.raises(ValueError):normalize({'radius_um':9.,'delta_n':.01})


def test_single_fem_uses_requested_parameters():
    a=run_job({**DEFAULTS,'core_elements':160,'clad_elements':640})
    b=run_job({**DEFAULTS,'core_elements':160,'clad_elements':640,'radius_um':3.8})
    ma={x['key']:x for x in a['metrics']};mb={x['key']:x for x in b['metrics']}
    assert ma['neff']['error']<2e-7
    assert ma['aeff']['error']<.01
    assert abs(ma['aeff']['value']-mb['aeff']['value'])>1
    assert a['details']['elements']==800
    assert a['plots']['mode']['image'].startswith('data:image/png;base64,')
    assert a['inputs']['radius_um']==4.1 and b['inputs']['radius_um']==3.8


def test_hcf_mesh_convergence_and_wall_interpretation():
    result=run_job({**DEFAULTS,'model':'hcf','hcf_elements':320},'convergence')
    assert len(result['rows'])==3
    assert result['rows'][-1]['neff_error']<result['rows'][0]['neff_error']/10
    assert result['inputs']['hcf_elements']==320
    with pytest.raises(ValueError,match='실제 모델 경계'):run_job({'model':'hcf'},'domain')


def test_mcf_fem_splitting_and_units():
    result=run_job({**DEFAULTS,'model':'mcf','outer_um':32.,'mesh_size_um':1.})
    m={x['key']:x for x in result['metrics']}
    assert m['split']['value']>0 and m['Lc']['error']<.01
    assert m['kappa']['value']==pytest.approx(np.pi/(2*m['Lc']['value']*1e-3),rel=1e-10)
    assert result['details']['dofs']>1000


def test_spectral_changes_are_recomputed():
    r=run_job({**DEFAULTS,'core_elements':40,'clad_elements':160,'sweep_points':3},'sweep')
    assert len(r['rows'])==3
    assert r['rows'][0]['x']==1530 and r['rows'][-1]['x']==1625
    assert r['rows'][0]['aeff']!=r['rows'][-1]['aeff']
    assert 'sweep' in r['plots']


def test_catalog_and_default_saved_example():
    catalog=json.loads((ROOT/'catalog.json').read_text(encoding='utf-8'));sample=json.loads((ROOT/'sample_result.json').read_text(encoding='utf-8'))
    assert catalog['defaults']==DEFAULTS==sample['inputs']
    assert all(normalize(p['values'])==p['values'] for p in PRESETS)
    assert sample['action']=='solve'


@pytest.fixture(scope='module')
def server():
    s=WorkbenchServer(('127.0.0.1',0));thread=threading.Thread(target=s.serve_forever,daemon=True);thread.start()
    yield 'http://127.0.0.1:'+str(s.server_address[1])
    s.shutdown();s.server_close();thread.join()


def request(server,path,data=None,headers=None):
    req=urllib.request.Request(server+path,data=json.dumps(data).encode() if data is not None else None,headers=headers or ({'Content-Type':'application/json'} if data is not None else {}))
    return urllib.request.urlopen(req,timeout=30)


def test_http_health_static_and_input_rejection(server):
    assert json.load(request(server,'/api/status'))['ready']
    assert b'LS Wave Optics' in request(server,'/').read()
    with pytest.raises(urllib.error.HTTPError) as e:request(server,'/api/jobs',{'inputs':{'model':'mcf','pitch_um':3}})
    assert e.value.code==400
    with pytest.raises(urllib.error.HTTPError) as e:request(server,'/api/jobs',{'inputs':DEFAULTS},headers={'Content-Type':'application/json','Origin':'https://example.org'})
    assert e.value.code==403
    with pytest.raises(urllib.error.HTTPError):request(server,'/../engine.py')


def test_http_job_runs_real_fem_in_worker(server):
    job=json.load(request(server,'/api/jobs',{'action':'solve','inputs':{**DEFAULTS,'model':'hcf','hcf_elements':160}}))
    for _ in range(150):
        r=json.load(request(server,'/api/jobs/'+job['id']))
        if r['status'] in ['completed','failed']:break
        time.sleep(.1)
    assert r['status']=='completed',r
    neff=r['result']['metrics'][0]
    assert neff['error']<4e-9
    assert r['result']['inputs']['model']=='hcf'


def test_fem_petermann_ii_against_bessel_gradient_integral():
    from engine import compute_one,make_fiber
    from fiber_modes import lp01_analytic
    from scipy.integrate import quad
    from scipy.special import j0,j1,k0,k1
    c=normalize({});f=make_fiber(c);a=lp01_analytic(f);u,w=a['u'],a['w']
    I=quad(lambda t:(j0(u*t)/j0(u))**2*t,0,1)[0]+quad(lambda t:(k0(w*t)/k0(w))**2*t,1,np.inf)[0]
    J=quad(lambda t:(u*j1(u*t)/j0(u))**2*t,0,1)[0]+quad(lambda t:(w*k1(w*t)/k0(w))**2*t,1,np.inf)[0]
    reference=2*f.radius_um*np.sqrt(2*I/J)
    actual=next(x['value'] for x in compute_one(c,plots=False,dispersion=False)['metrics'] if x['key']=='mfd_pii')
    assert actual==pytest.approx(reference,rel=1e-4)
