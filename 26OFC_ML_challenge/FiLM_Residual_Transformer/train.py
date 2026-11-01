#!/usr/bin/env python3
"""Train/resume/evaluate Residual-FiLM EDFA on official OFC 2026 95-channel CSVs.

Examples:
    python train.py --data-root /content/ofc-ml-challenge-data-code --output results
    python train.py --data-root /content/ofc-ml-challenge-data-code --pretrain-cosmos \
       --cosmos-max-rows 4000 --pretrain-epochs 15 --epochs 200 --scratch-control

Only Train features/labels are used to fit, tune or select the model.
Optional TestGround access is a *post-selection audit*, never part of training.
"""
from __future__ import annotations

import argparse
import copy
import json
import random
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import torch
from torch import nn
from torch.utils.data import DataLoader, TensorDataset

from data import (N_CHANNELS, TARGET_PREFIX, load_csv, load_official_pair, to_arrays,
                  fit_condition_stats, split_groups, subset_indices)
from model import ResidualFiLMEDFA, SHBAwareCurvatureLoss


def seed_all(seed: int):
    random.seed(seed)
    np.random.seed(seed)
    torch.manual_seed(seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed_all(seed)


def official_score(y_true: np.ndarray, y_pred: np.ndarray, active_mask: np.ndarray) -> dict:
    """Match the OFC original competitor pipeline's per-row loaded-channel score.

    This is a NumPy *evaluation metric*, not part of backpropagation.
    """
    active = active_mask.astype(bool)
    n_on = active.sum(axis=1)
    valid = n_on > 0
    if not np.any(valid):
        raise ValueError('No active channels to evaluate')
    diff = np.abs(y_true - y_pred)
    safe = np.where(active, diff, 0.0)
    row_mae = safe.sum(axis=1)[valid] / n_on[valid]
    vmask = active[valid]
    vdiff = safe[valid]
    row_std = np.sqrt(np.where(vmask, (vdiff - row_mae[:,None])**2, 0).sum(axis=1) / n_on[valid])
    flat = diff[active]
    mae, std = float(row_mae.mean()), float(row_std.mean())
    p95, maximum = float(np.quantile(flat,.95)), float(flat.max())
    tail95 = max(0., p95-mae-0.5)
    tailmax = max(0., maximum-p95-0.7)
    return {'score':mae+0.3*tail95+0.1*tailmax+0.15*std,
            'mae':mae,'std':std,'p95':p95,'max_error':maximum,
            'tail95':tail95, 'tailmax':tailmax,
            'samples':int(valid.sum()),'active_channels':int(active.sum())}


def loader(arrays: dict, indices: np.ndarray, batch_size: int, shuffle: bool):
    keys = ('pin_dbm','active_mask','aging','target_gain','target_tilt','gain_db')
    ts = [torch.as_tensor(np.ascontiguousarray(arrays[k][indices]), dtype=torch.float32) for k in keys]
    return DataLoader(TensorDataset(*ts), batch_size=batch_size,
                      shuffle=shuffle, num_workers=0)


def forward_batch(model, batch, device):
    pin, mask, age, gain, tilt, truth = [a.to(device) for a in batch]
    pred, components = model(pin, mask, age, gain, tilt)
    return pred, components, truth, mask


@torch.no_grad()
def predict_arrays(model, arrays: dict, device, batch_size: int = 512):
    model.eval()
    n = len(arrays['aging'])
    output = []
    for start in range(0,n,batch_size):
        end = min(start+batch_size,n)
        keys = ('pin_dbm','active_mask','aging','target_gain','target_tilt')
        inp = [torch.from_numpy(np.ascontiguousarray(arrays[k][start:end])).to(device) for k in keys]
        pred, _ = model(*inp)
        output.append(pred.cpu().numpy())
    return np.concatenate(output) if output else np.empty((0,N_CHANNELS), dtype=np.float32)


def evaluate_indexed(model, arrays, indices, device):
    subset={k:v[indices] for k,v in arrays.items()}
    prediction = predict_arrays(model, subset, device)
    return official_score(subset['gain_db'], prediction, subset['active_mask']), prediction


def run_stage(model, dataset, train_idx, val_idx, device, args, stage, epochs, lr, patience, out):
    criterion = SHBAwareCurvatureLoss(lambda_curvature=args.curvature,
                                      tolerance_db=args.curvature_tolerance)
    optimizer = torch.optim.AdamW(model.parameters(), lr=lr, weight_decay=args.weight_decay)
    scheduler = torch.optim.lr_scheduler.ReduceLROnPlateau(optimizer, mode='min', factor=0.65, patience=10)
    train_loader = loader(dataset, train_idx, args.batch_size, shuffle=True)
    history=[]; best=float('inf'); best_state=None; bad=0
    for epoch in range(1, epochs+1):
        model.train(); train_loss=[]; mse_loss=[]; curv_loss=[]
        for batch in train_loader:
            optimizer.zero_grad(set_to_none=True)
            pred,_,truth,mask=forward_batch(model,batch,device)
            loss,parts=criterion(pred,truth,mask)
            if not torch.isfinite(loss):
                raise RuntimeError(f'Non-finite loss at stage={stage} epoch={epoch}')
            loss.backward()
            nn.utils.clip_grad_norm_(model.parameters(), max_norm=2.0)
            optimizer.step()
            train_loss.append(float(loss.detach()))
            mse_loss.append(float(parts['mse']))
            curv_loss.append(float(parts['curvature']))
        metrics,_=evaluate_indexed(model,dataset,val_idx,device)
        val=metrics['score']; scheduler.step(val)
        history.append({'stage':stage,'epoch':epoch,'train_loss':float(np.mean(train_loss)),
                        'train_mse':float(np.mean(mse_loss)),
                        'train_curvature':float(np.mean(curv_loss)),
                        'val_score':val,'val_mae':metrics['mae'],
                        'lr':optimizer.param_groups[0]['lr']})
        if val < best-1e-5:
            best=val;bad=0
            best_state={k:v.detach().cpu().clone() for k,v in model.state_dict().items()}
        else:
            bad+=1
        if epoch==1 or epoch%10==0 or epoch==epochs:
            print(f'[{stage}] epoch={epoch:3d}, loss={np.mean(train_loss):.5f}, val={val:.5f}, best={best:.5f}',flush=True)
        if bad>=patience:
            print(f'[{stage}] early stopping at {epoch}',flush=True)
            break
    if best_state is None:
        raise RuntimeError('No finite validation checkpoint found')
    model.load_state_dict(best_state)
    # CRITICAL: report the restored best state, not the last-epoch scores.
    best_metrics,_=evaluate_indexed(model,dataset,val_idx,device)
    pd.DataFrame(history).to_csv(out/f'{stage}_history.csv',index=False)
    checkpoint={'model_state':best_state,'model_kwargs':model.config(),
                'stage':stage,'best_metrics':best_metrics, 'seed':args.seed}
    torch.save(checkpoint,out/f'{stage}_best.pt')
    return best_metrics


def per_category(df, y, pred, mask):
    cat=next((c for c in df if c.lower()=='category'),None)
    result={'ALL':official_score(y,pred,mask)}
    if cat:
        vals=df[cat].fillna('unknown').astype(str).str.lower().to_numpy()
        for name in sorted(np.unique(vals)):
            choose=vals==name
            result[name]=official_score(y[choose],pred[choose],mask[choose])
    return result


def save_validation(out, label, model, df, arrays, indices, device):
    sub={k:v[indices] for k,v in arrays.items()}
    p=predict_arrays(model,sub,device)
    m=sub['active_mask'].astype(bool)
    cats=df.iloc[indices]['Category'].astype(str).to_numpy() if 'Category' in df else ['all']*len(indices)
    ids=df.iloc[indices]['ID'].astype(str).to_numpy() if 'ID' in df else indices
    summary=per_category(df.iloc[indices],sub['gain_db'],p,sub['active_mask'])
    (out/f'{label}_metrics.json').write_text(json.dumps(summary,indent=2,ensure_ascii=False))
    rows=[]
    for j in range(len(indices)):
        on=m[j]
        if not on.any(): continue
        e=np.abs((p[j]-sub['gain_db'][j])[on])
        rows.append({'ID':ids[j], 'Category':cats[j],
                     'target_gain':float(sub['target_gain'][j]),
                     'target_tilt':float(sub['target_tilt'][j]),
                     'n_active':int(on.sum()), 'mae_db':float(e.mean()),
                     'max_abs_error_db':float(e.max())})
    pd.DataFrame(rows).to_csv(out/f'{label}_per_sample.csv',index=False)
    print(f'[{label}] best restored checkpoint:',json.dumps(summary['ALL']))
    return summary


def make_submission(test_df, prediction: np.ndarray, active_mask: np.ndarray, output: Path):
    columns=[f'{TARGET_PREFIX}{i:02d}' for i in range(N_CHANNELS)]
    sub=pd.DataFrame(np.where(active_mask>0.5,prediction,-18.),columns=columns)
    id_col=next((c for c in test_df if c.lower()=='id'),None)
    if id_col is None:
        raise ValueError('Test features must provide ID for official submission schema')
    sub.insert(0,'ID',test_df[id_col].values)
    sub.to_csv(output,index=False,float_format='%.6f')
    return sub


def parse_args(argv=None):
    ap=argparse.ArgumentParser(description='OFC 2026 Residual + FiLM Transformer with SHB-aware curvature regularization')
    ap.add_argument('--data-root',type=Path,default=Path('ofc-ml-challenge-data-code'))
    ap.add_argument('--output',type=Path,default=Path('film_residual_results'))
    ap.add_argument('--epochs',type=int,default=200)
    ap.add_argument('--patience',type=int,default=35)
    ap.add_argument('--lr',type=float,default=0.002,help='Recent Train training/fine-tuning LR (0.002 is previous transfer control)')
    ap.add_argument('--batch-size',type=int,default=64)
    ap.add_argument('--hidden',type=int,default=64)
    ap.add_argument('--layers',type=int,default=2)
    ap.add_argument('--heads',type=int,default=4)
    ap.add_argument('--curvature',type=float,default=0.02)
    ap.add_argument('--curvature-tolerance',type=float,default=0.05)
    ap.add_argument('--weight-decay',type=float,default=2e-4)
    ap.add_argument('--val-frac',type=float,default=0.2)
    ap.add_argument('--age-column',type=str,default=None,help='explicit numeric age column in FEATURES, not Category')
    ap.add_argument('--pretrain-cosmos',action='store_true')
    ap.add_argument('--cosmos-max-rows',type=int,default=0,help='0=entire available COSMOS Train split')
    ap.add_argument('--pretrain-epochs',type=int,default=60)
    ap.add_argument('--pretrain-patience',type=int,default=15)
    ap.add_argument('--pretrain-lr',type=float,default=0.001)
    ap.add_argument('--scratch-control',action='store_true')
    ap.add_argument('--postcontest-testground',action='store_true',help='Evaluate after model selection, without affecting training')
    ap.add_argument('--audit-checkpoint',type=Path,default=None,help='Skip training; evaluate a frozen .pt checkpoint against TestGround')
    ap.add_argument('--cpu',action='store_true')
    ap.add_argument('--seed',type=int,default=2026)
    args=ap.parse_args(argv)
    if args.epochs<1 or args.pretrain_epochs<1 or args.batch_size<1:
        ap.error('epochs and batch size must be >= 1')
    if args.hidden%args.heads:
        ap.error('hidden must be divisible by heads')
    if not 0 < args.val_frac < 1 or args.curvature<0:
        ap.error('val-frac must be in (0,1), curvature >=0')
    return args


def run(args):
    seed_all(args.seed)
    out=args.output; out.mkdir(parents=True,exist_ok=True)
    device=torch.device('cpu' if args.cpu or not torch.cuda.is_available() else 'cuda')
    if args.audit_checkpoint is not None:
        # Strict post-training audit: the original checkpoint is never updated.
        checkpoint = torch.load(args.audit_checkpoint, map_location='cpu', weights_only=True)
        frozen = ResidualFiLMEDFA(**checkpoint['model_kwargs']).to(device)
        frozen.load_state_dict(checkpoint['model_state'])
        gp=args.data_root/'Features'/'TestGround'
        gx=load_csv(gp/'test_features.csv')
        gl=load_csv(gp/'test_labels.csv')
        ground=to_arrays(gx,gl,args.age_column)
        yp=predict_arrays(frozen,ground,device)
        result=per_category(gx,ground['gain_db'],yp,ground['active_mask'])
        (out/'frozen_checkpoint_testground_audit.json').write_text(
            json.dumps(result,indent=2,ensure_ascii=False))
        print('FROZEN postcontest TestGround audit:', json.dumps(result,indent=2))
        return result
    train_df,recent=load_official_pair(args.data_root,'train',args.age_column)
    if args.age_column is None and not any(c in train_df for c in ('aging_hours','aging_days','age_hours','age_days','EDFA_age_hours','device_age_hours','DUT_age','age')):
        warnings.warn('No measured numerical aging feature detected. Aging condition set to 0; never infer it from Category.')
    it,iv=split_groups(train_df,recent,args.val_frac,args.seed)
    print(f'Device={device} | Train={len(it)} | Validation={len(iv)} | active-val-channels={int(recent["active_mask"][iv].sum())}',flush=True)
    def build(mean,std):
        return ResidualFiLMEDFA(mean,std,hidden_dim=args.hidden,
                                n_heads=args.heads,n_layers=args.layers).to(device)
    stats=fit_condition_stats(recent,it)
    stages={}
    if args.pretrain_cosmos:
        cosmos_df,cosmos=load_official_pair(args.data_root,'COSMOS',args.age_column)
        ct,cv=split_groups(cosmos_df,cosmos,0.08,args.seed)
        ct=subset_indices(ct,args.cosmos_max_rows,args.seed)
        cv=subset_indices(cv,3000,args.seed+1)
        print(f'COSMOS pretraining {len(ct)} samples, validation {len(cv)}',flush=True)
        # Keep FiLM conditioning in the SAME source-coordinate system during fine-tuning.
        mean,std=fit_condition_stats(cosmos,ct)
        model=build(mean,std)
        stages['cosmos_pretrain']=run_stage(model,cosmos,ct,cv,device,args,'cosmos_pretrain',
                                             args.pretrain_epochs,args.pretrain_lr,args.pretrain_patience,out)
    else:
        model=build(*stats)
    stages['recent_best']=run_stage(model,recent,it,iv,device,args,'recent',
                                    args.epochs,args.lr,args.patience,out)
    stages['recent_by_category']=save_validation(out,'recent',model,train_df,recent,iv,device)
    if args.scratch_control and args.pretrain_cosmos:
        seed_all(args.seed)
        scratch=build(*stats)
        stages['scratch_best']=run_stage(scratch,recent,it,iv,device,args,'scratch',
                                         args.epochs,args.lr,args.patience,out)
        stages['scratch_by_category']=save_validation(out,'scratch',scratch,train_df,recent,iv,device)
    test_path=args.data_root/'Features'/'Test'/'test_features.csv'
    if test_path.exists():
        tdf=load_csv(test_path)
        test_arrays=to_arrays(tdf,None,args.age_column)
        prediction=predict_arrays(model,test_arrays,device)
        sub=make_submission(tdf,prediction,test_arrays['active_mask'],out/'submission.csv')
        print('Submission rows:',len(sub))
    if args.postcontest_testground:
        gp=args.data_root/'Features'/'TestGround'
        gx=load_csv(gp/'test_features.csv')
        gl=load_csv(gp/'test_labels.csv')
        ground=to_arrays(gx,gl,args.age_column)
        p=predict_arrays(model,ground,device)
        audit=per_category(gx,ground['gain_db'],p,ground['active_mask'])
        (out/'testground_postcontest_audit.json').write_text(json.dumps(audit,indent=2))
        stages['postcontest_testground']=audit
    (out/'comparison.json').write_text(json.dumps(stages,indent=2,ensure_ascii=False))
    pd.DataFrame([{'stage':name,**value} for name,value in stages.items()
                  if isinstance(value,dict) and 'score' in value]).to_csv(out/'summary.csv',index=False)
    print('Finished:',out.resolve(),flush=True)
    return stages


if __name__=='__main__':
    run(parse_args())