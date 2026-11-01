"""Small CPU integration test. All generated observations here are SYNTHETIC, NOT OFC scores."""
from __future__ import annotations
import json
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
import torch

from data import to_arrays, choose_age_column, fit_condition_stats
from model import ResidualFiLMEDFA, FiLMTransformerBlock, SHBAwareCurvatureLoss
from train import parse_args, run, official_score


def fake_data(n=54, seed=11):
    rng=np.random.default_rng(seed)
    pos=np.linspace(-1,1,95)
    row={'ID':np.arange(n),
         'Category':np.array(['aging','shb','unseen']*(n//3)+['unseen']*(n%3)),
         'target_gain':rng.choice([16.,18.,20.],n),
         'target_gain_tilt':rng.choice([-1.,0.,1.],n),
         'EDFA_input_power_total':-8.+rng.normal(size=n),
         'EDFA_output_power_total':11.+rng.normal(size=n)}
    mask=np.zeros((n,95),dtype=np.float32)
    pin=-26 + rng.normal(0,2,(n,95))
    for j in range(n):
        if j%6==0:mask[j]=1
        elif j%6==1:mask[j,5:60]=1
        elif j%6==2:mask[j,::2]=1
        elif j%6==3:mask[j,30:50]=1
        elif j%6==4:mask[j,[4,78]]=1
        else:mask[j,40]=1
    gain=(row['target_gain'][:,None]+.5*row['target_gain_tilt'][:,None]*pos[None,:]
          +0.10*np.cos(pos*8)[None,:]
          +rng.normal(0,.09,(n,95)))
    # SHB measured local hole, not a model artifact.
    gain[1]+=(-0.9*np.exp(-.5*((np.arange(95)-42)/2.)**2))
    for i in range(95):
        row[f'EDFA_input_spectra_{i:02d}']=pin[:,i]
        row[f'DUT_WSS_activated_channel_index_{i:02d}']=mask[:,i]
    X=pd.DataFrame(row)
    Y=pd.DataFrame({'ID':X.ID})
    for i in range(95):Y[f'calculated_gain_spectra_{i:02d}']=gain[:,i].astype(np.float32)
    return X,Y


def test_model_and_loss():
    torch.manual_seed(2026)
    model=ResidualFiLMEDFA([0.,18.,0.],[1.,2.,1.],hidden_dim=32,n_heads=4,n_layers=2)
    pin=torch.randn(4,95)-20.
    mask=torch.ones(4,95)
    mask[1,:]=0.;mask[1,8]=1.;mask[1,70]=1.
    truth=torch.randn(4,95)+18.
    truth[mask==0]=float('nan')
    pred,parts=model(pin,mask,torch.zeros(4),torch.full((4,),18.),torch.zeros(4))
    assert pred.shape==(4,95)
    loss_fn=SHBAwareCurvatureLoss()
    loss,components=loss_fn(pred,truth,mask)
    assert torch.isfinite(loss)
    assert torch.isfinite(components['curvature'])
    loss.backward()
    assert model.residual_head.weight.grad is not None
    assert 'cond_mean' in model.state_dict()
    # Formula preserves target SHB curvature when prediction equals truth.
    good_truth=torch.where(mask.bool(),truth,torch.zeros_like(truth))
    zero_loss, parts0 = loss_fn(good_truth,truth,mask)
    assert abs(zero_loss.item())<1e-8
    assert parts0['curvature'].item()<1e-8
    assert choose_age_column(pd.DataFrame({'Category':['aging']})) is None
    print('MODEL, FILM, LOSS, SHB-AWARE MASK: OK')


def test_official_pipeline(tmp: Path):
    X,Y=fake_data()
    root=tmp/'dataset'
    for sub in ['Train','Test']:(root/'Features'/sub).mkdir(parents=True)
    X.to_csv(root/'Features/Train/train_features.csv',index=False)
    Y.to_csv(root/'Features/Train/train_labels.csv',index=False)
    X.iloc[:7].to_csv(root/'Features/Test/test_features.csv',index=False)
    X.iloc[:42].to_csv(root/'Features/Train/COSMOS_features.csv',index=False)
    Y.iloc[:42].to_csv(root/'Features/Train/COSMOS_labels.csv',index=False)
    A=to_arrays(X,Y)
    assert np.all(A['aging']==0) # category is NOT age / no leakage
    assert A['pin_dbm'].shape==(len(X),95)
    assert A['active_mask'].shape==(len(X),95)
    assert official_score(A['gain_db'],A['gain_db'],A['active_mask'])['score'] < 1e-12
    output=tmp/'results'
    args=parse_args(['--data-root',str(root),'--output',str(output),'--cpu',
                     '--hidden','32','--heads','4','--layers','2',
                     '--epochs','2','--pretrain-epochs','2','--patience','2',
                     '--pretrain-cosmos','--cosmos-max-rows','20','--scratch-control',
                     '--batch-size','16','--lr','0.002'])
    metrics=run(args)
    assert all(k in metrics for k in ('cosmos_pretrain','recent_best','scratch_best'))
    for file in ['cosmos_pretrain_best.pt','recent_best.pt','scratch_best.pt',
                 'recent_history.csv','recent_metrics.json','scratch_metrics.json',
                 'submission.csv','comparison.json']:
        assert (output/file).exists(),file
    checkpoint=torch.load(output/'recent_best.pt',map_location='cpu',weights_only=True)
    recovered=ResidualFiLMEDFA(**checkpoint['model_kwargs'])
    recovered.load_state_dict(checkpoint['model_state'])
    assert recovered(torch.randn(2,95),torch.ones(2,95),torch.zeros(2),
                     torch.full((2,),18.),torch.zeros(2))[0].shape==(2,95)
    submission=pd.read_csv(output/'submission.csv')
    assert submission.shape==(7,96)
    assert 'ID' in submission
    print('END-TO-END OFFICIAL CSV SCHEMA, COSMOS TRANSFER, TRAIN, CHECKPOINT, SUBMISSION: OK')
    print('TEST ONLY (SYNTHETIC DATA):',json.dumps({k:round(v['score'],4) for k,v in metrics.items() if 'score' in v}))


if __name__=='__main__':
    test_model_and_loss()
    with tempfile.TemporaryDirectory() as folder:
        test_official_pipeline(Path(folder))