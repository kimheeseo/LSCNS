#!/usr/bin/env python3
"""Two-stage training for the existing PI-EDFA Transformer; no winner weights."""
import argparse, json, sys, copy, time, hashlib
from pathlib import Path
import numpy as np
import pandas as pd
import torch
from sklearn.model_selection import GroupShuffleSplit, train_test_split

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "PI_EDFA_Transformer"))
import transformer_edfa as core

def split_indices(df, seed, frac):
    mask=core.extract_mask(df)
    groups=core.get_groups(df,mask)
    if len(set(groups))>5:
        train,val=next(GroupShuffleSplit(n_splits=1,test_size=frac,random_state=seed).split(df,groups=groups))
        return train,val,"group"
    train,val=train_test_split(np.arange(len(df)),test_size=frac,random_state=seed)
    return train,val,"random_fallback"

def load_pair(fp,lp):
    x=core.read_table(Path(fp)).reset_index(drop=True)
    y=core.read_table(Path(lp)).reset_index(drop=True)
    if len(x)!=len(y): raise ValueError("Feature and label row counts differ")
    cols=core.ordered_cols(y,core.PREFIXES["gain"])
    vals=y[cols].apply(pd.to_numeric,errors="coerce").to_numpy(dtype=np.float32)
    mask=core.extract_mask(x)
    invalid=(mask.sum(axis=1)==0)|(~np.isfinite(vals)&(mask>0)).any(axis=1)
    print(f"제외: 활성 채널 없음 또는 활성 채널 정답 비유한 {invalid.sum()}행",flush=True)
    keep=~invalid
    x=x.loc[keep].reset_index(drop=True); vals=np.where(mask[keep]>0,vals[keep],0)
    if len(x)<10: raise ValueError("Insufficient valid rows")
    return x,vals

def main():
    p=argparse.ArgumentParser()
    for name in ["cosmos-features","cosmos-labels","train-features","train-labels","test-features"]:
        p.add_argument("--"+name,required=True)
    p.add_argument("--output",default="improved_results")
    p.add_argument("--seed",type=int,default=2026)
    p.add_argument("--val-frac",type=float,default=.2)
    p.add_argument("--pretrain-epochs",type=int,default=100)
    p.add_argument("--finetune-epochs",type=int,default=100)
    p.add_argument("--batch-size",type=int,default=64)
    p.add_argument("--hidden",type=int,default=64)
    p.add_argument("--patience",type=int,default=20)
    p.add_argument("--pretrain-lr",type=float,default=.001)
    p.add_argument("--finetune-lr",type=float,default=.0002)
    p.add_argument("--cpu",action="store_true")
    args=p.parse_args()
    if not 0<args.val_frac<1: raise ValueError("val-frac must be between 0 and 1")
    out=Path(args.output);out.mkdir(parents=True,exist_ok=True)
    core.seed_everything(args.seed)
    cx,cy=load_pair(args.cosmos_features,args.cosmos_labels)
    tx,ty=load_pair(args.train_features,args.train_labels)
    ti,vi,kind=split_indices(tx,args.seed,args.val_frac)
    # Block exact feature duplicates of any competition sample, including validation.
    feature_cols=[c for c in tx.columns if c in cx.columns and c.lower() not in ("id","category")]
    if not feature_cols: raise ValueError("No common columns for duplicate check")
    # pandas stable hashing compares data, excluding identifiers.
    train_hash=set(pd.util.hash_pandas_object(tx[feature_cols],index=False).to_numpy())
    chash=pd.util.hash_pandas_object(cx[feature_cols],index=False).to_numpy()
    keep=~np.isin(chash,list(train_hash))
    print(f"대회 데이터와 동일 입력의 COSMOS {int((~keep).sum())}행 제외; 유사 샘플은 별도 검토 필요")
    cx=cx.loc[keep].reset_index(drop=True);cy=cy[keep]
    if len(cx)<10: raise ValueError("Insufficient COSMOS rows after deduplication")
    ci,cvi,ckind=split_indices(cx,args.seed,args.val_frac)
    # A single frozen scaler fitted on COSMOS train + competition train, never validation/test.
    processor=core.EDFAFeatures().fit(pd.concat([cx.iloc[ci],tx.iloc[ti]],ignore_index=True))
    cf=processor.transform(cx);tf=processor.transform(tx)
    cm=core.extract_mask(cx);tm=core.extract_mask(tx)
    device="cuda" if torch.cuda.is_available() and not args.cpu else "cpu"
    net=core.PhysicalSpectralNet(global_dim=tf[0].shape[1],hidden=args.hidden).to(device)
    history=[]
    def stage(name,feat,labels,masks,train,val,epochs,lr):
        arrays=feat+[labels,masks]
        tr=core.make_loader(arrays,train,args.batch_size,True)
        va=core.make_loader(arrays,val,args.batch_size,False)
        opt=torch.optim.AdamW(net.parameters(),lr=lr,weight_decay=2e-4)
        sched=torch.optim.lr_scheduler.ReduceLROnPlateau(opt,mode="min",patience=10,factor=.6)
        best=float("inf");state=None;bad=0
        for ep in range(1,epochs+1):
            net.train();losses=[]
            for batch in tr:
                b=[z.to(device) for z in batch]
                opt.zero_grad(set_to_none=True)
                pred,components=net(*b[:6])
                loss=core.loss_fn(pred,b[6],b[7],components)
                if not torch.isfinite(loss): raise RuntimeError("Nonfinite loss")
                loss.backward();torch.nn.utils.clip_grad_norm_(net.parameters(),3.);opt.step()
                losses.append(loss.item())
            net.eval();pred=[];actual=[];ms=[]
            with torch.no_grad():
                for batch in va:
                    b=[z.to(device) for z in batch]
                    yhat,_=net(*b[:6]);pred.append(yhat.cpu().numpy());actual.append(batch[6].numpy());ms.append(batch[7].numpy())
            result=core.official_score(np.concatenate(actual),np.concatenate(pred),np.concatenate(ms))
            score=float(result["score"]);sched.step(score)
            history.append(dict(stage=name,epoch=ep,train_loss=float(np.mean(losses)),val_score=score,val_mae=float(result["mae"])))
            print(f"{name} {ep}/{epochs} val_score={score:.6f} val_MAE={result['mae']:.6f}",flush=True)
            if score<best-1e-5: best=score;state=copy.deepcopy(net.state_dict());bad=0
            else: bad+=1
            if bad>=args.patience: break
        if state is None: raise ValueError("No training epochs; set epochs >= 1")
        net.load_state_dict(state)
        return best
    started=time.perf_counter()
    pre=stage("COSMOS_pretrain",cf,cy,cm,ci,cvi,args.pretrain_epochs,args.pretrain_lr)
    ft=stage("competition_finetune",tf,ty,tm,ti,vi,args.finetune_epochs,args.finetune_lr)
    val_pred=core.predict_model(net,[x[vi] for x in tf],device)
    metrics=core.official_score(ty[vi],val_pred,tm[vi])
    test=core.read_table(Path(args.test_features))
    testfeat=processor.transform(test)
    t=time.perf_counter();prediction=core.predict_model(net,testfeat,device);infer=time.perf_counter()-t
    core.make_submission(test,prediction,out/"submission.csv")
    pd.DataFrame(history).to_csv(out/"history.csv",index=False)
    np.savez(out/"split_indices.npz",cosmos_train=ci,cosmos_val=cvi,competition_train=ti,competition_val=vi)
    torch.save({"model":net.state_dict(),"scaler_mean":processor.scaler.mean_,"scaler_scale":processor.scaler.scale_,"hidden":args.hidden,"seed":args.seed},out/"model.pt")
    result={"competition_validation":metrics,"cosmos_validation_score":pre,"competition_split":kind,"cosmos_split":ckind,"parameters":sum(p.numel() for p in net.parameters()),"test_rows":len(test),"inference_seconds":infer,"elapsed_seconds":time.perf_counter()-started,"note":"Internal validation only. Not official hidden leaderboard. Cross-dataset near-duplicates are not ruled out."}
    (out/"metrics.json").write_text(json.dumps(result,indent=2,ensure_ascii=False,default=float),encoding="utf-8")
    print(json.dumps(result,indent=2,ensure_ascii=False,default=float))
if __name__=="__main__": main()
