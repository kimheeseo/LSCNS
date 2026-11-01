#!/usr/bin/env python3
"""Reproduce the independent 20-dB OFC 2026 measured-data sanity benchmark.

IMPORTANT: This is a physics-informed spectral RIDGE SURROGATE, NOT the previous
PI-EDFA PyTorch convolutional model. Data source is the official repository's
measurement JSON (not Git-LFS CSV pointers). The test ground truth is used only
for one final score and not for training, normalization, or model selection.

Usage (Python 3.10+):
  pip install numpy pandas requests matplotlib
  python official_json_subset_benchmark.py --output results_official_subset
"""
from __future__ import annotations
import argparse, json
from pathlib import Path
import numpy as np
import pandas as pd
import requests
import matplotlib.pyplot as plt

ORG_REPO = 'ofc-ml-challenge/ofc-ml-challenge-data-code'
BASE = f'https://raw.githubusercontent.com/{ORG_REPO}/main/dataset/ML_challenge_admin'
DEVICES = ('rdm1-lg1', 'rdm2-lg1', 'rdm3-co1', 'rdm4-co1', 'rdm5-co1', 'rdm6-co1')
TRAIN = ([f'aging/{part}' for part in (
    'booster_rdm1-co1_fix_train.json', 'booster_rdm2-co1_fix_train.json',
    'preamp_rdm1-co1_fix_train.json', 'preamp_rdm2-co1_fix_train.json')]
    + [f'shb/preamp_{d}_random_train.json' for d in ('rdm1-co1', 'rdm2-co1')]
    + [f'unseen/tilt_booster_{d}_fix_train.json' for d in DEVICES])
TEST = [f'unseen/gain_booster_{d}_fix_testGround.json' for d in DEVICES]


def fetch(path: str, dataset: str, cachedir: Path) -> dict:
    f = cachedir / dataset / path
    if not f.exists():
        f.parent.mkdir(parents=True, exist_ok=True)
        url = f'{BASE}/{dataset}/{path}'
        print('Downloading', url, flush=True)
        resp = requests.get(url, timeout=90)
        resp.raise_for_status()
        # The JSON files are ordinary Git blobs; do not confuse them with LFS pointers.
        blob = resp.json()
        f.write_text(json.dumps(blob), encoding='utf-8')
    return json.loads(f.read_text(encoding='utf-8'))


def convert(path: str, dataset: str, cache: Path) -> list[dict]:
    src = fetch(path, dataset, cache)
    preamp = 'preamp' in path
    category = path.split('/')[0]
    import re
    device = re.search(r'rdm\d+-(?:co|lg)\d+', path).group()
    items = []
    for m in src['measurement_data']:
        cfg = m['roadm_dut_preamp_info'] if preamp else m['roadm_dut_edfa_info']
        inp = m['roadm_dut_preamp_input_power_spectra'] if preamp else m['roadm_dut_wss_output_power_spectra']
        out = m['roadm_dut_wss_input_power_spectra'] if preamp else m['roadm_dut_booster_output']
        loaded = (m['roadm_flatten_wss_active_channel_index'] if preamp else
                  m['roadm_dut_wss_active_channel_index'])
        # JSON object keys are the numbered channels 1..95.
        p = np.array([float(inp[str(i)]) for i in range(1, 96)])
        o = np.array([float(out[str(i)]) for i in range(1, 96)])
        mask = np.zeros(95, dtype=bool)
        mask[np.array(loaded, dtype=int)-1] = True
        items.append({'category': category, 'device': device,
                      'g': float(cfg['target_gain']),
                      't': float(cfg['target_gain_tilt']),
                      'pin': float(cfg['input_power']),
                      'pout': float(cfg['output_power']),
                      'p': p, 'y': o-p, 'mask': mask})
    return items


def physical_features(s: dict) -> np.ndarray:
    x = np.arange(95, dtype=float)/47. - 1.
    mask, p = s['mask'], s['p']
    count = max(1, int(mask.sum()))
    avg = p[mask].mean() if mask.any() else 0.
    powermw = np.where(mask, 10**(np.clip(p, -80, 20)/10), 0.)
    share = powermw / max(1., powermw.sum()) if powermw.sum()==0 else powermw / powermw.sum()
    dist = np.arange(95, dtype=float)[:, None] - np.arange(95, dtype=float)[None, :]
    kernels = [1/(1+(dist/w)**2) @ share for w in (2,7,20)]
    kernel_means = [v[mask].mean() if mask.any() else 0. for v in kernels]
    dev = np.clip((p-avg)/10, -5, 5)
    log_share = np.log1p(95*share)
    occ = count/95
    delta = s['pout']-s['pin']-s['g']
    pin = (s['pin']+10)/20
    gain = s['g']-18
    cols = [np.ones(95), x, x*x, x**3,
            np.full(95, gain), gain*x,
            np.full(95, s['t']), s['t']*x, s['t']*x*x,
            np.full(95, pin), pin*x, pin*x*x,
            np.full(95, delta), delta*x, np.full(95, occ), occ*x,
            dev, dev*x, log_share, log_share*x]
    cols += [(v-mn)*25 for v,mn in zip(kernels,kernel_means)]
    cols += [np.exp(-.5*((x-(-1+i/4))/.31)**2) for i in range(9)]
    return np.column_stack(cols)


def baseline(s):
    return s['g']+.5*s['t']*(np.arange(95)/47-1)


def ridge_fit(rows: list[dict], alpha: float = 2.):
    F = [physical_features(s) for s in rows]
    dim = F[0].shape[1]
    X, y, weights = [], [], []
    for s,f in zip(rows,F):
        on = s['mask']
        X.append(f[on]);y.append((s['y']-baseline(s))[on]);
        weights.append(np.full(int(on.sum()), 1/max(1,int(on.sum()))))
    X=np.concatenate(X); y=np.concatenate(y); w=np.concatenate(weights)
    totw=w.sum();mu=(X*w[:,None]).sum(axis=0)/totw
    std=np.maximum(.001,np.sqrt(np.maximum(0,(X*X*w[:,None]).sum(axis=0)/totw-mu**2)))
    mu[0]=0;std[0]=1
    Z=(X-mu)/std
    gram=Z.T@(w[:,None]*Z)
    regularizer=np.eye(dim)*alpha;regularizer[0,0]=1e-5
    beta=np.linalg.solve(gram+regularizer, Z.T@(w*y))
    return beta,mu,std


def prediction(s,coef):
    beta,mu,std=coef
    return baseline(s)+((physical_features(s)-mu)/std)@beta


def official_score(rows, outputs):
    byrow=[]; flat=[]
    for s,pred in zip(rows,outputs):
        err=np.abs((s['y']-pred)[s['mask']])
        if len(err):
            byrow.append((err.mean(),err.std()))
            flat.append(err)
    flat=np.concatenate(flat)
    mae=float(np.mean([x[0] for x in byrow]));std=float(np.mean([x[1] for x in byrow]))
    p95=float(np.quantile(flat,.95));mx=float(flat.max())
    score=mae+.15*std+.3*max(0,p95-mae-.5)+.1*max(0,mx-p95-.7)
    return {'score':score,'mae':mae,'std':std,'p95':p95,'max_error':mx,
            'rmse':float(np.sqrt(np.mean(flat**2))),
            'samples':len(byrow),'loaded_points':len(flat),
            'within_0p2':float(np.mean(flat<=.2))}


def run(output, cachedir):
    train=[s for p in TRAIN for s in convert(p,'Train',cachedir)]
    test=[s for p in TEST for s in convert(p,'TestGround',cachedir)]
    assert len(train)==530 and len(test)==306, (len(train),len(test))
    assert {s['g'] for s in train}=={18} and {s['g'] for s in test}=={20}
    coef=ridge_fit(train)
    b=[baseline(s) for s in test]
    pred=[prediction(s,coef) for s in test]
    outcome={'setup':{'train_rows':len(train),'test_rows':len(test),
                      'features':32,'ridge_alpha':2,
                      'validation':'Only 20-dB gain / fixed-loading TestGround, post-contest data',
                      'note':'This ridge surrogate is NOT the PI-EDFA neural model'},
             'baseline':official_score(test,b),
             'physics_ridge':official_score(test,pred)}
    devrecords=[]
    for dev in DEVICES:
        idx=[k for k,s in enumerate(test) if s['device']==dev]
        a=[test[k] for k in idx]
        v={'device':dev,'baseline':official_score(a,[b[k] for k in idx]),
           'physics_ridge':official_score(a,[pred[k] for k in idx])}
        devrecords.append(v)
    outcome['by_device']=devrecords
    output.mkdir(parents=True,exist_ok=True)
    (output/'metrics.json').write_text(json.dumps(outcome,indent=2),encoding='utf-8')
    pd.DataFrame([{'device':v['device'],'baseline_score':v['baseline']['score'],
                   'ridge_score':v['physics_ridge']['score'],
                   'baseline_mae':v['baseline']['mae'],
                   'ridge_mae':v['physics_ridge']['mae']} for v in devrecords]).to_csv(
                       output/'device_scores.csv',index=False)
    fig,ax=plt.subplots(figsize=(8,4))
    ax.bar(['Gain/Tilt only','Physics-guided ridge'],[outcome['baseline']['score'],outcome['physics_ridge']['score']])
    ax.set_ylabel('OFC 2026 score (lower is better)')
    ax.set_title('Official measured TestGround / Unseen 20 dB (6 devices)')
    fig.tight_layout();fig.savefig(output/'score_comparison.png',dpi=150);plt.close(fig)
    print(json.dumps(outcome,indent=2))


if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--output',default='results_official_subset')
    ap.add_argument('--cache',default='.ofc_raw_json_cache');a=ap.parse_args()
    run(Path(a.output),Path(a.cache))
