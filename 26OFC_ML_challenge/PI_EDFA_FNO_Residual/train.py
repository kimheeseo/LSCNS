#!/usr/bin/env python3
"""Train PI-EDFA baseline and a frozen-PI + FNO residual on OFC 2026 CSVs.

TestGround is loaded only when explicitly requested after model selection.
"""
from __future__ import annotations
import argparse, copy, json
from pathlib import Path
import numpy as np
import pandas as pd
import torch
from torch.utils.data import DataLoader, TensorDataset
from model import PIEDFAFNO, pi

def parse_args():
    a=argparse.ArgumentParser()
    a.add_argument('--data-root',required=True)
    a.add_argument('--output',default='fno_results')
    a.add_argument('--seed',type=int,default=2026)
    a.add_argument('--val-frac',type=float,default=.2)
    a.add_argument('--pi-epochs',type=int,default=180)
    a.add_argument('--fno-epochs',type=int,default=180)
    a.add_argument('--pretrain-epochs',type=int,default=0,
                   help='Optional COSMOS pretrain of PI backbone only (0=disabled)')
    a.add_argument('--cosmos-max-rows',type=int,default=4000,
                   help='0 means all COSMOS rows if enabled')
    a.add_argument('--pi-lr',type=float,default=.002)
    a.add_argument('--fno-lr',type=float,default=.0005)
    a.add_argument('--batch-size',type=int,default=64)
    a.add_argument('--patience',type=int,default=30)
    a.add_argument('--pi-hidden',type=int,default=96)
    a.add_argument('--fno-width',type=int,default=48)
    a.add_argument('--fno-modes',type=int,default=16)
    a.add_argument('--fno-depth',type=int,default=3)
    a.add_argument('--cpu',action='store_true')
    a.add_argument('--evaluate-testground',action='store_true',
                   help='Post-selection audit only; do not tune on TestGround')
    args=a.parse_args()
    if not 0<args.val_frac<1:a.error('val-frac must be in (0,1)')
    if min(args.pi_epochs,args.fno_epochs)<1:a.error('epochs must be positive')
    if args.pretrain_epochs<0 or args.cosmos_max_rows<0:a.error('invalid COSMOS args')
    return args

def load_pair(root,stem):
    folder=root/'Features'/'Train'
    x=pi.read_table(folder/(stem+'_features.csv')).reset_index(drop=True)
    labels=pi.read_table(folder/(stem+'_labels.csv')).reset_index(drop=True)
    if len(x)!=len(labels) or ('ID' in x and 'ID' in labels and not x.ID.equals(labels.ID)):
        raise ValueError('Features/labels ID mismatch: '+stem)
    mask=pi.extract_mask(x).astype(bool)
    y=labels[pi.ordered_cols(labels,pi.PREFIXES['gain'])].apply(
        pd.to_numeric,errors='coerce').to_numpy(np.float32)
    if not np.isfinite(y[mask]).all():raise ValueError('Nonfinite active labels')
    return x,np.where(mask,y,0.).astype(np.float32)

def loader(features,y,idx,batch,shuffle=False):
    arrays=[np.ascontiguousarray(a[idx]) for a in features]
    arrays+=[np.ascontiguousarray(y[idx]),np.ascontiguousarray(features[2][idx])]
    return DataLoader(TensorDataset(*(torch.as_tensor(a,dtype=torch.float32)
                                      for a in arrays)),batch_size=batch,
                      shuffle=shuffle,num_workers=0)

@torch.no_grad()
def predict(net,features,device,batch_size=256):
    net.eval()
    outputs=[]
    for i in range(0,len(features[0]),batch_size):
        batch=[torch.as_tensor(v[i:i+batch_size],dtype=torch.float32,device=device)
               for v in features]
        y,_=net(*batch)
        outputs.append(y.detach().cpu().numpy())
    return np.concatenate(outputs) if outputs else np.zeros((0,pi.N_CH),np.float32)

def metrics_by_category(x,y,pred,mask):
    result={'ALL':pi.official_score(y,pred,mask)}
    col=next((c for c in x.columns if c.lower()=='category'),None)
    if col:
        vals=x[col].fillna('unknown').astype(str).str.lower().to_numpy()
        for cat in sorted(set(vals)):
            ix=vals==cat
            result[cat]=pi.official_score(y[ix],pred[ix],mask[ix])
    return result

def fit_stage(net,features,y,tr,va,args,out,stage,epochs,lr,device):
    train=loader(features,y,tr,args.batch_size,True)
    valid=loader(features,y,va,max(args.batch_size,256))
    params=[p for p in net.parameters() if p.requires_grad]
    opt=torch.optim.AdamW(params,lr=lr,weight_decay=2e-4)
    sch=torch.optim.lr_scheduler.ReduceLROnPlateau(opt,'min',patience=8,factor=.6)
    best=float('inf');bad=0;history=[];state=None
    for ep in range(1,epochs+1):
        net.train()
        if hasattr(net,'physics') and not any(p.requires_grad for p in net.physics.parameters()):
            net.physics.eval()
        losses=[]
        for sample in train:
            b=[v.to(device) for v in sample]
            opt.zero_grad(set_to_none=True)
            pred,comp=net(*b[:6])
            loss=pi.loss_fn(pred,b[6],b[7],comp)
            if not torch.isfinite(loss):raise ValueError('Nonfinite loss: '+stage)
            loss.backward()
            torch.nn.utils.clip_grad_norm_(params,3.)
            opt.step()
            losses.append(float(loss.item()))
        pred=[];actual=[];masks=[]
        net.eval()
        with torch.no_grad():
            for sample in valid:
                b=[v.to(device) for v in sample]
                yp,_=net(*b[:6])
                pred.append(yp.cpu().numpy())
                actual.append(sample[6].numpy())
                masks.append(sample[7].numpy())
        score=pi.official_score(np.concatenate(actual),np.concatenate(pred),
                                np.concatenate(masks))['score']
        sch.step(score)
        history.append({'stage':stage,'epoch':ep,'train_loss':float(np.mean(losses)),
                        'val_score':score,'lr':opt.param_groups[0]['lr']})
        if score<best-1e-6:
            best=score;bad=0
            state=copy.deepcopy({k:v.detach().cpu() for k,v in net.state_dict().items()})
        else:bad+=1
        if ep==1 or ep%10==0 or ep==epochs:
            print(f'{stage} {ep}/{epochs}: score={score:.6f}, best={best:.6f}',flush=True)
        if bad>=args.patience:break
    net.load_state_dict(state)
    pd.DataFrame(history).to_csv(out/(stage+'_history.csv'),index=False)
    return best

def evaluate_testground(root,out,net,processor,device):
    # TestGround never affects optimizer, feature scalers, or checkpoint selection.
    folder=root/'Features'/'TestGround'
    x=pi.read_table(folder/'test_features.csv').reset_index(drop=True)
    labels=pi.read_table(folder/'test_labels.csv').reset_index(drop=True)
    if len(x)!=len(labels) or ('ID' in x and 'ID' in labels and not x.ID.equals(labels.ID)):
        raise ValueError('TestGround ID mismatch')
    mask=pi.extract_mask(x).astype(bool)
    y=labels[pi.ordered_cols(labels,pi.PREFIXES['gain'])].apply(
        pd.to_numeric,errors='coerce').to_numpy(np.float32)
    if not np.isfinite(y[mask]).all():raise ValueError('Nonfinite active TestGround label')
    y=np.where(mask,y,0.).astype(np.float32)
    features=processor.transform(x)
    pred=predict(net,features,device)
    baseline=predict(net.physics,features,device)
    results={'PI_EDFA':metrics_by_category(x,y,baseline,mask),
             'PI_EDFA_FNO':metrics_by_category(x,y,pred,mask)}
    (out/'testground_comparison.json').write_text(json.dumps(results,indent=2))
    print('POST-SELECTION TESTGROUND (NOT Kaggle hidden leaderboard):',
          json.dumps(results,indent=2),flush=True)
    return results

def main():
    args=parse_args()
    torch.set_num_threads(min(torch.get_num_threads(),4))
    pi.seed_everything(args.seed)
    root=Path(args.data_root);out=Path(args.output);out.mkdir(parents=True,exist_ok=True)
    x,y=load_pair(root,'train')
    mask=pi.extract_mask(x)
    groups=pi.get_groups(x,mask)
    idx=np.arange(len(x))
    if len(np.unique(groups))>1:
        from sklearn.model_selection import GroupShuffleSplit
        tr,va=next(GroupShuffleSplit(n_splits=1,test_size=args.val_frac,
                                      random_state=args.seed).split(idx,groups=groups))
    else:
        from sklearn.model_selection import train_test_split
        tr,va=train_test_split(idx,test_size=args.val_frac,random_state=args.seed)
    processor=pi.EDFAFeatures().fit(x.iloc[tr])
    features=processor.transform(x)
    device='cpu' if args.cpu or not torch.cuda.is_available() else 'cuda'
    net=PIEDFAFNO(len(processor.scaler.mean_),args.pi_hidden,args.fno_width,
                  args.fno_modes,args.fno_depth).to(device)
    if args.pretrain_epochs:
        cx,cy=load_pair(root,'COSMOS')
        ids=np.arange(len(cx))
        if args.cosmos_max_rows and len(ids)>args.cosmos_max_rows:
            ids=np.sort(np.random.default_rng(args.seed).choice(
                ids,args.cosmos_max_rows,replace=False))
        order=np.random.default_rng(args.seed+2).permutation(ids)
        if len(order)<5:raise ValueError('COSMOS requires at least five rows')
        n_val=max(1,int(len(order)*.1));ct,cv=order[n_val:],order[:n_val]
        fit_stage(net.physics,processor.transform(cx),cy,ct,cv,args,out,
                  'cosmos_pi_pretrain',args.pretrain_epochs,args.pi_lr,device)
    fit_stage(net.physics,features,y,tr,va,args,out,'recent_pi',
              args.pi_epochs,args.pi_lr,device)
    net.freeze_physics()
    baseline=predict(net.physics,[f[va] for f in features],device)
    best=fit_stage(net,features,y,tr,va,args,out,'recent_fno',
                   args.fno_epochs,args.fno_lr,device)
    improved=predict(net,[f[va] for f in features],device)
    xv=x.iloc[va].reset_index(drop=True);yv=y[va];mv=mask[va]
    result={'protocol':'recent Train grouped holdout, no TestGround model selection',
            'seed':args.seed,'train_rows':len(tr),'validation_rows':len(va),
            'cosmos_pretrain_epochs':args.pretrain_epochs,
            'PI_EDFA':metrics_by_category(xv,yv,baseline,mv),
            'PI_EDFA_FNO':metrics_by_category(xv,yv,improved,mv),
            'fno_best_val_score':best}
    (out/'validation_comparison.json').write_text(json.dumps(result,indent=2))
    print('TRAIN HOLDOUT:',json.dumps(result,indent=2),flush=True)
    torch.save({'state_dict':net.state_dict(),
                'config':{'global_dim':len(processor.scaler.mean_),
                          'pi_hidden':args.pi_hidden,'width':args.fno_width,
                          'modes':args.fno_modes,'depth':args.fno_depth},
                'scaler_mean':torch.as_tensor(processor.scaler.mean_,dtype=torch.float64),
                'scaler_scale':torch.as_tensor(processor.scaler.scale_,dtype=torch.float64),
                'seed':args.seed},out/'best_pi_edfa_fno.pt')
    path=root/'Features/Test/test_features.csv'
    if path.exists():
        tx=pi.read_table(path)
        pi.make_submission(tx,predict(net,processor.transform(tx),device),
                           out/'submission.csv')
        print('submission.csv created:',out/'submission.csv',flush=True)
    if args.evaluate_testground:
        evaluate_testground(root,out,net,processor,device)
    print('Finished:',out.resolve(),flush=True)

if __name__=='__main__':main()
