#!/usr/bin/env python3
"""PI-EDFA Transformer: gain/tilt baseline plus channel attention residual.

Inputs: target gain and tilt, optical total powers, 95-channel input spectrum, WSS mask.
Outputs: 95 gain values. No winner architecture or pretrained winner weights used.

Dependencies: torch, numpy, pandas, scikit-learn, matplotlib.
"""
from __future__ import annotations
import argparse, json, random, re, warnings
from pathlib import Path
import numpy as np
import pandas as pd
import torch
from torch import nn
import torch.nn.functional as F
from torch.utils.data import DataLoader, TensorDataset
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import GroupShuffleSplit, train_test_split
import matplotlib.pyplot as plt

N_CH = 95
SEED = 2026
PREFIXES = {
    "input": "EDFA_input_spectra_",
    "mask": "DUT_WSS_activated_channel_index_",
    "gain": "calculated_gain_spectra_",
}
SCALARS = ["target_gain", "target_gain_tilt", "EDFA_input_power_total", "EDFA_output_power_total"]


def seed_everything(seed=SEED):
    random.seed(seed); np.random.seed(seed); torch.manual_seed(seed)
    if torch.cuda.is_available(): torch.cuda.manual_seed_all(seed)


def read_table(path: Path) -> pd.DataFrame:
    if not path.exists(): raise FileNotFoundError(f"File not found: {path}")
    with path.open("rb") as f: head = f.read(128)
    if head.startswith(b"version https://git-lfs.github.com/spec/v1"):
        raise RuntimeError(f"Git LFS pointer instead of CSV: {path}\n"
                           "Run 'git lfs pull --include=Features/Train/*.csv,Features/Test/*.csv,Features/TestGround/*.csv' ")
    return pd.read_csv(path)


def ordered_cols(df: pd.DataFrame, prefix: str) -> list[str]:
    r = re.compile(r"^" + re.escape(prefix) + r"(\d+)$", re.I)
    found = [(int(m.group(1)), col) for col in df.columns if (m := r.match(col))]
    found.sort()
    if len(found) != N_CH or [i for i,_ in found] != list(range(N_CH)):
        raise ValueError(f"Expected 95 channel columns for '{prefix}', got {len(found)} in {list(df.columns)[:12]}")
    return [col for _,col in found]


def extract_mask(df):
    return (df[ordered_cols(df, PREFIXES['mask'])].apply(pd.to_numeric, errors='coerce')
            .fillna(0).to_numpy(dtype=np.float32) > 0.5).astype(np.float32)


class EDFAFeatures:
    """Standardize only GLOBAL features; spectral powers stay in physical units."""
    def __init__(self): self.scaler = StandardScaler()

    @staticmethod
    def raw(df):
        missing = [s for s in SCALARS if s not in df]
        if missing: raise ValueError(f"Missing scalar columns: {missing}")
        p_dbm = df[ordered_cols(df, PREFIXES['input'])].apply(pd.to_numeric, errors='coerce')
        p_dbm = p_dbm.to_numpy(dtype=np.float32)
        m = extract_mask(df)
        p_dbm = np.nan_to_num(p_dbm, nan=-80., posinf=15., neginf=-80.).clip(-80., 25.)
        # 'input power' is in dBm, not watts: convert to mW before summing.
        p_mw = np.power(10., p_dbm/10.).astype(np.float32) * m
        power_sum = p_mw.sum(axis=1, keepdims=True)
        share = p_mw / np.maximum(power_sum, 1e-12)
        count = np.maximum(m.sum(axis=1, keepdims=True), 1.)
        spectral_mean = (p_dbm*m).sum(axis=1,keepdims=True) / count
        spectral_std = np.sqrt((((p_dbm-spectral_mean)*m)**2).sum(axis=1,keepdims=True)/count)
        x = np.linspace(-1.,1.,N_CH, dtype=np.float32)[None,:]
        occupancy = m.mean(axis=1,keepdims=True)
        centroid = (share*x).sum(axis=1,keepdims=True)
        spectral_dev = np.clip((p_dbm-spectral_mean)/10., -6., 6.)*m
        scalars = np.column_stack([
            df[s].apply(pd.to_numeric, errors='coerce').fillna(0).to_numpy(dtype=np.float32)
            for s in SCALARS
        ]).astype(np.float32)
        delta_power = (scalars[:,3]-scalars[:,2])[:,None]
        x_global = np.concatenate([scalars, occupancy, spectral_mean, spectral_std,
                                   delta_power, centroid, np.log1p(power_sum)], axis=1)
        # Channels: occupancy, normalized measured spectral shape, fraction of input
        # mW, and wavelength position. No labels or metadata leakage.
        x_chan = np.stack([m, spectral_dev, np.log1p(95.*share),
                           np.broadcast_to(x,m.shape)], axis=1).astype(np.float32)
        return x_global, x_chan, m, share, scalars[:,0], scalars[:,1]

    def fit(self, df): self.scaler.fit(self.raw(df)[0]); return self

    def transform(self, df):
        global_raw, chan, mask, share, target, tilt = self.raw(df)
        global_norm = self.scaler.transform(global_raw).astype(np.float32)
        return [global_norm, chan, mask, share, target, tilt]


class PhysicalSpectralNet(nn.Module):
    """Gain/tilt reference plus a compact global channel-attention residual.

    All 95 positions are retained as keys, including unloaded positions: their
    mask and zero power are explicit inputs. This avoids all-masked attention.
    Only loaded output channels contribute to the training and evaluation loss.
    """
    def __init__(self, global_dim=10, hidden=64, n_basis=13):
        super().__init__()
        self.register_buffer('x', torch.linspace(-1., 1., N_CH))
        self.global_net = nn.Sequential(nn.Linear(global_dim, hidden), nn.SiLU(),
                                        nn.Linear(hidden, hidden))
        self.channel_embed = nn.Linear(4, hidden)
        self.blocks = nn.TransformerEncoder(
            nn.TransformerEncoderLayer(d_model=hidden, nhead=4,
                                       dim_feedforward=2*hidden, dropout=.08,
                                       activation='gelu', batch_first=True,
                                       norm_first=True), num_layers=2,
            enable_nested_tensor=False)
        self.residual_head = nn.Sequential(nn.LayerNorm(hidden), nn.Linear(hidden, 1))
        self.tilt_scale = nn.Parameter(torch.tensor(.5))
        nn.init.zeros_(self.residual_head[-1].weight)
        nn.init.zeros_(self.residual_head[-1].bias)

    def forward(self, global_x, channel_x, mask, share, target_gain, target_tilt):
        base = target_gain[:, None] + self.tilt_scale * target_tilt[:, None] * self.x[None, :]
        tokens = self.channel_embed(channel_x.transpose(1, 2))
        tokens = tokens + self.global_net(global_x)[:, None, :]
        correction = self.residual_head(self.blocks(tokens)).squeeze(-1)
        return base + correction, {'base': base, 'local_correction': correction}


def official_score(y_true, y_pred, mask):
    """OFC 2026/Kaggle formula, using only loaded channels (row-wise MAE & STD)."""
    mask = np.asarray(mask).astype(bool)
    error = np.abs(np.asarray(y_true)-np.asarray(y_pred))
    loaded = mask.sum(axis=1)
    valid = loaded>0
    if not valid.any(): return {'score':float('nan'),'mae':float('nan')}
    row_mae = (error*mask).sum(axis=1)[valid]/loaded[valid]
    e=error[valid]; m=mask[valid]
    row_std=np.sqrt((((e-row_mae[:,None])**2)*m).sum(axis=1)/loaded[valid])
    flat=e[m]
    mae=float(row_mae.mean()); std=float(row_std.mean());p95=float(np.quantile(flat,.95));emax=float(flat.max())
    tail95=max(0.,p95-mae-0.5)
    tailmax=max(0.,emax-p95-0.7)
    score=mae+0.3*tail95+0.1*tailmax+0.15*std
    return dict(score=float(score),mae=mae,std=std,p95=p95,max_error=emax,
                penalty_p95=tail95,penalty_max=tailmax,samples=int(valid.sum()),
                loaded_points=int(flat.size))


def extended_metrics(y_true, y_pred, mask):
    """Supplementary absolute-error diagnostics, never evaluated on off channels."""
    on = np.asarray(mask).astype(bool)
    residual = np.asarray(y_pred) - np.asarray(y_true)
    error = residual[on]
    if not error.size:
        return {'rmse': float('nan'), 'mean_bias': float('nan'),
                'within_0p1_db': float('nan'), 'within_0p2_db': float('nan')}
    ae = np.abs(error)
    return {'rmse': float(np.sqrt(np.mean(error ** 2))),
            'mean_bias': float(np.mean(error)),
            'median_absolute_error': float(np.median(ae)),
            'p99_absolute_error': float(np.quantile(ae, .99)),
            'within_0p1_db': float(np.mean(ae <= .1)),
            'within_0p2_db': float(np.mean(ae <= .2)),
            'within_0p5_db': float(np.mean(ae <= .5))}


def evaluate_predictions(outdir, prefix, metadata, y_true, y_pred, mask, baseline=None):
    """Persist repeatable performance evaluation for one untouched evaluation split.

    Arrays are per measured EDFA setting: [samples, 95 channels]; only masked-on
    channels are scored. Saved CSVs let users inspect errors without retraining.
    """
    outdir = Path(outdir)
    y_true = np.asarray(y_true, dtype=np.float64)
    y_pred = np.asarray(y_pred, dtype=np.float64)
    mask = np.asarray(mask).astype(bool)
    if y_true.shape != y_pred.shape or y_true.shape != mask.shape or y_true.ndim != 2:
        raise ValueError('Expected matching [N,95] actual, predicted and mask arrays.')
    if y_true.shape[1] != N_CH or len(metadata) != len(y_true):
        raise ValueError('Evaluation row counts or 95-channel shape mismatch.')
    base = (np.asarray(baseline, dtype=np.float64)
            if baseline is not None else None)
    if base is not None and base.shape != y_true.shape:
        raise ValueError('Baseline array must be [N,95].')
    active_n = mask.sum(axis=1)
    residual = y_pred - y_true
    ae = np.abs(residual)
    ids = metadata['ID'].to_numpy() if 'ID' in metadata.columns else np.arange(len(metadata))
    catcol = next((c for c in metadata.columns if c.lower()=='category'), None)
    categories = (metadata[catcol].fillna('unknown').astype(str).to_numpy()
                  if catcol is not None else np.array(['all']*len(metadata)))
    row_data=[]
    for i in range(len(y_true)):
        on=mask[i]
        err=residual[i,on]
        abs_err=np.abs(err)
        row={'original_index':str(metadata.index[i]), 'ID':ids[i],
             'Category':categories[i],
             'target_gain':metadata.iloc[i].get('target_gain',np.nan),
             'target_gain_tilt':metadata.iloc[i].get('target_gain_tilt',np.nan),
             'loaded_channels':int(on.sum()),
             'mae_db':float(abs_err.mean()) if on.any() else np.nan,
             'rmse_db':float(np.sqrt((err**2).mean())) if on.any() else np.nan,
             'bias_db':float(err.mean()) if on.any() else np.nan,
             'max_error_db':float(abs_err.max()) if on.any() else np.nan}
        if base is not None:
            row['baseline_mae_db']=(float(np.abs(base[i,on]-y_true[i,on]).mean())
                                    if on.any() else np.nan)
        row_data.append(row)
    row_df=pd.DataFrame(row_data)
    row_df.to_csv(outdir/f'{prefix}_per_sample.csv', index=False)
    chan_data=[]
    for c in range(N_CH):
        e=residual[mask[:,c],c]
        b=(base[mask[:,c],c]-y_true[mask[:,c],c]) if base is not None else None
        row={'channel':c,'loaded_samples':len(e),
             'mae_db':float(np.abs(e).mean()) if len(e) else np.nan,
             'rmse_db':float(np.sqrt((e**2).mean())) if len(e) else np.nan,
             'bias_db':float(e.mean()) if len(e) else np.nan,
             'p95_error_db':float(np.quantile(np.abs(e),.95)) if len(e) else np.nan,
             'within_0p2_db':float(np.mean(np.abs(e)<=.2)) if len(e) else np.nan}
        if b is not None:
            row['baseline_mae_db']=float(np.abs(b).mean()) if len(e) else np.nan
        chan_data.append(row)
    channel_df=pd.DataFrame(chan_data)
    channel_df.to_csv(outdir/f'{prefix}_per_channel.csv',index=False)
    # One wide row per observation, including original IDs and precise channel masks.
    wide = pd.DataFrame({'ID':ids,'Category':categories})
    for tag,values in [('measured',y_true),('predicted',y_pred),('active',mask.astype(int))]:
        wide = pd.concat([wide, pd.DataFrame(values, columns=[f'{tag}_{c:02d}' for c in range(N_CH)])],axis=1)
    wide.to_csv(outdir/f'{prefix}_predictions.csv',index=False,float_format='%.6f')
    score=official_score(y_true,y_pred,mask)
    score.update(extended_metrics(y_true,y_pred,mask))
    bycat={}
    for cat in np.unique(categories):
        ix=categories==cat
        bycat[str(cat)]={**official_score(y_true[ix],y_pred[ix],mask[ix]),
                         **extended_metrics(y_true[ix],y_pred[ix],mask[ix])}
    (outdir/f'{prefix}_by_category.json').write_text(json.dumps(bycat,indent=2,ensure_ascii=False))
    pd.DataFrame([{'category':name,**details} for name,details in bycat.items()]).to_csv(
        outdir/f'{prefix}_by_category.csv',index=False)
    (outdir/f'{prefix}_extended_metrics.json').write_text(json.dumps(score,indent=2,ensure_ascii=False))
    if len(y_true):
        fig,axes=plt.subplots(1,2,figsize=(12,4.2))
        rng=np.random.default_rng(2026)
        scatter_e=residual[mask]
        scatter_y=y_true[mask]
        if len(scatter_e)>25000:
            idx=rng.choice(len(scatter_e),25000,replace=False)
            scatter_e,scatter_y=scatter_e[idx],scatter_y[idx]
        axes[0].scatter(scatter_y,scatter_y+scatter_e,s=5,alpha=.24)
        low=min(float(scatter_y.min()),float((scatter_y+scatter_e).min()))
        high=max(float(scatter_y.max()),float((scatter_y+scatter_e).max()))
        axes[0].plot([low,high],[low,high],'k--',lw=1)
        axes[0].set(xlabel='Measured gain (dB)',ylabel='Predicted gain (dB)',title='Measured vs predicted')
        axes[0].grid(alpha=.2)
        axes[1].plot(channel_df['channel'],channel_df['mae_db'],label='Physics-guided model')
        if base is not None:axes[1].plot(channel_df['channel'],channel_df['baseline_mae_db'],label='Gain/Tilt baseline',ls='--')
        axes[1].set(xlabel='Channel index (0-94)',ylabel='MAE (dB)',title='Per-channel absolute error')
        axes[1].grid(alpha=.2);axes[1].legend()
        fig.tight_layout();fig.savefig(outdir/f'{prefix}_performance.png',dpi=170);plt.close(fig)
        valid=row_df.dropna(subset=['mae_db']).sort_values('mae_db',ascending=False)
        selected=valid.head(min(3,len(valid))).index.to_list()
        if selected:
            fig,axes=plt.subplots(len(selected),1,figsize=(10,3*len(selected)),squeeze=False)
            for ax,ri in zip(axes[:,0],selected):
                ci=np.flatnonzero(mask[ri]);ax.plot(ci,y_true[ri,ci],'-o',ms=2,label='Measured')
                ax.plot(ci,y_pred[ri,ci],'-',label='Predicted')
                ax.set(title=f'High-error sample: ID {ids[ri]}, category {categories[ri]} | MAE={row_df.iloc[ri].mae_db:.3f} dB',
                       xlabel='Channel index',ylabel='Gain (dB)')
                ax.grid(alpha=.25);ax.legend()
            fig.tight_layout();fig.savefig(outdir/f'{prefix}_worst_samples.png',dpi=170);plt.close(fig)
    return score,bycat


def loss_fn(pred, y, mask, components=None):
    e=torch.abs(pred-y)
    n=mask.sum(dim=1).clamp_min(1.)
    per_row_mae=(e*mask).sum(dim=1)/n
    per_row_std=torch.sqrt((((e-per_row_mae[:,None])**2)*mask).sum(dim=1)/n+1e-8)
    # Differentiable approximation of the official score; small shape prior.
    base=per_row_mae.mean()+0.15*per_row_std.mean()
    active=e[mask>0.5]
    if active.numel():
        p95=torch.quantile(active,.95)
        base=base+0.3*F.relu(p95-per_row_mae.mean()-.5)
        base=base+0.04*F.relu(active.max()-p95-.7)
    if components is not None:
        # Mild inductive prior: local learned correction is kept small unless
        # enough data supports it; broad spectrum is modeled with smooth bases.
        base=base+1e-4*components['local_correction'].square().mean()
    return base


def get_groups(df, mask):
    """Keep repeated gain/tilt/device/channel-load setups together if present."""
    device_candidates = [c for c in df.columns if c.lower() in ('edfa_name','edfa_index','edfa_type','edfa_id','device')]
    device = df[device_candidates[0]].fillna('unknown').astype(str).values if device_candidates else ['device']*len(df)
    gain=pd.to_numeric(df['target_gain'], errors='coerce').fillna(0).to_numpy()
    tilt=pd.to_numeric(df['target_gain_tilt'],errors='coerce').fillna(0).to_numpy()
    groups=[]
    for k in range(len(df)):
        packed=np.packbits(mask[k].astype(np.uint8)).tobytes().hex()
        groups.append(f'{device[k]}:{gain[k]:.2f}:{tilt[k]:.2f}:{packed}')
    return np.asarray(groups)


def make_loader(tensors,indices,batch_size,shuffle):
    a=[torch.from_numpy(t[indices].astype('float32')) for t in tensors]
    return DataLoader(TensorDataset(*a),batch_size=batch_size,shuffle=shuffle,num_workers=0)


@torch.no_grad()
def predict_model(model, features, device='cpu', batch=512):
    model.eval()
    res=[]
    arrays=[torch.from_numpy(x.astype('float32')) for x in features]
    for i in range(0,len(arrays[0]),batch):
        xs=[a[i:i+batch].to(device) for a in arrays]
        p,_=model(*xs);res.append(p.cpu().numpy())
    return np.concatenate(res) if res else np.empty((0,N_CH),np.float32)


def train_model(train_df,y_df,args):
    seed_everything(args.seed)
    train_mask=extract_mask(train_df)
    label_cols=ordered_cols(y_df,PREFIXES['gain'])
    y=y_df[label_cols].apply(pd.to_numeric,errors='coerce').to_numpy(dtype=np.float32)
    y=np.nan_to_num(y,nan=0.,posinf=0.,neginf=0.)
    y=np.where(train_mask>0.5,y,0.)
    if len(train_df)!=len(y): raise ValueError('Train input and label row counts differ')
    group=get_groups(train_df,train_mask)
    if len(set(group))>5:
        split=GroupShuffleSplit(n_splits=1,test_size=args.val_frac,random_state=args.seed)
        idx_train,idx_val=next(split.split(train_df,groups=group))
    else:
        idx_train,idx_val=train_test_split(np.arange(len(train_df)),test_size=args.val_frac,random_state=args.seed)
    if not len(idx_train) or not len(idx_val): raise ValueError('Train/val split is empty')
    # Fit scaler ONLY on training split: no test or holdout leakage.
    feature_maker=EDFAFeatures().fit(train_df.iloc[idx_train])
    feats=feature_maker.transform(train_df)
    all_arrays=feats+[y,train_mask]
    train_loader=make_loader(all_arrays,idx_train,args.batch_size,True)
    val_loader=make_loader(all_arrays,idx_val,args.batch_size,False)
    device='cuda' if torch.cuda.is_available() and not args.cpu else 'cpu'
    net=PhysicalSpectralNet(global_dim=feats[0].shape[1],hidden=args.hidden).to(device)
    opt=torch.optim.AdamW(net.parameters(),lr=args.lr,weight_decay=2e-4)
    sched=torch.optim.lr_scheduler.ReduceLROnPlateau(opt,mode='min',patience=10,factor=.6)
    best=float('inf');best_state=None;bad=0;history=[]
    for ep in range(1,args.epochs+1):
        net.train();tr=[]
        for tensors in train_loader:
            data=[x.to(device) for x in tensors]; x=data[:6]; lab,mask=data[6:]
            opt.zero_grad(set_to_none=True)
            prediction,components=net(*x)
            loss=loss_fn(prediction,lab,mask,components)
            if not torch.isfinite(loss): raise RuntimeError('Non-finite loss. Check units and data.')
            loss.backward();torch.nn.utils.clip_grad_norm_(net.parameters(),max_norm=3.)
            opt.step();tr.append(float(loss.detach()))
        net.eval();pred=[];actual=[];masks=[]
        with torch.no_grad():
            for batch in val_loader:
                arrays=[v.to(device) for v in batch]
                yy,_=net(*arrays[:6])
                pred.append(yy.cpu().numpy());actual.append(batch[6].numpy());masks.append(batch[7].numpy())
        result=official_score(np.concatenate(actual),np.concatenate(pred),np.concatenate(masks))
        v=result['score'];sched.step(v)
        history.append({'epoch':ep,'train_loss':float(np.mean(tr)),'val_score':v,'val_mae':result['mae']})
        if v<best-1e-5:
            best=v;bad=0
            best_state={k:vv.cpu().clone() for k,vv in net.state_dict().items()}
        else:bad+=1
        if ep==1 or ep%10==0 or ep==args.epochs:
            print(f'Epoch {ep:4d}  train-loss={np.mean(tr):.5f}  val-score={v:.5f}  best={best:.5f}')
        if bad>=args.patience:
            print(f'Early stopping at epoch {ep}, best validation score {best:.6f}')
            break
    net.load_state_dict(best_state)
    v_pred=predict_model(net,[x[idx_val] for x in feats],device)
    metrics=official_score(y[idx_val],v_pred,train_mask[idx_val]);
    print('Strict holdout metrics:',json.dumps(metrics,ensure_ascii=False,indent=2))
    return net,feature_maker,metrics,pd.DataFrame(history),idx_val,y[idx_val],v_pred,train_mask[idx_val]


def make_submission(test_df,pred,save_path):
    if pred.shape!=(len(test_df),N_CH):raise ValueError('Prediction must be N x 95')
    mask=extract_mask(test_df)
    cols=[f'{PREFIXES["gain"]}{i:02d}' for i in range(N_CH)]
    submission=pd.DataFrame(np.where(mask>0.5,pred,-18.0),columns=cols)
    idcol=next((c for c in test_df if c.lower()=='id'),test_df.columns[0])
    # competition provided IDs must be preserved exactly.
    submission.insert(0,'ID',test_df[idcol].to_numpy())
    submission.to_csv(save_path,index=False,float_format='%.6f')
    return submission


def plot_diagnostics(outdir,history,y,pred,mask):
    plt.figure(figsize=(7,4));plt.plot(history.epoch,history.train_loss,label='train surrogate loss')
    plt.plot(history.epoch,history.val_score,label='validation official score')
    plt.xlabel('Epoch');plt.ylabel('Error (dB)');plt.grid(alpha=.25);plt.legend();plt.tight_layout()
    plt.savefig(outdir/'training_curve.png',dpi=180);plt.close()
    loaded=mask.astype(bool);err=(pred-y)[loaded]
    plt.figure(figsize=(7,4));plt.hist(err,bins=50);plt.xlabel('Gain prediction error (dB)')
    plt.ylabel('Loaded channels');plt.grid(alpha=.25);plt.tight_layout()
    plt.savefig(outdir/'validation_error_histogram.png',dpi=180);plt.close()
    row=int(np.argmax(mask.sum(axis=1)))
    plt.figure(figsize=(8,4));idx=np.where(loaded[row])[0]
    plt.plot(idx,y[row,idx],'-o',ms=2.5,label='Measured')
    plt.plot(idx,pred[row,idx],'-',label='Physics-guided model')
    plt.xlabel('Channel index (0-94)');plt.ylabel('Gain (dB)');plt.grid(alpha=.25)
    plt.legend();plt.tight_layout();plt.savefig(outdir/'validation_spectrum.png',dpi=180);plt.close()


def checkpoint_to_model(checkpoint_file, cpu=False):
    """Load a trusted checkpoint written by this script, without any retraining."""
    ckpt=torch.load(checkpoint_file,map_location='cpu',weights_only=True)
    config=ckpt['config']
    model=PhysicalSpectralNet(global_dim=config['global_dim'],hidden=config['hidden'])
    model.load_state_dict(ckpt['state_dict'])
    device='cuda' if torch.cuda.is_available() and not cpu else 'cpu'
    model=model.to(device).eval()
    processor=EDFAFeatures()
    scaler=processor.scaler
    scaler.mean_=torch.as_tensor(ckpt['input_scaler_mean']).cpu().numpy().astype('float64')
    scaler.scale_=torch.as_tensor(ckpt['input_scaler_scale']).cpu().numpy().astype('float64')
    scaler.var_=scaler.scale_**2
    scaler.n_features_in_=len(scaler.mean_)
    scaler.n_samples_seen_=1
    return model,processor


def evaluate_testground(root,out,net,processor):
    """Archived ground truth: evaluation only, never used to fit or select a model."""
    ground=root/'Features/TestGround'
    gx=read_table(ground/'test_features.csv')
    gy=read_table(ground/'test_labels.csv')
    if len(gx)!=len(gy):raise ValueError('TestGround input/label lengths differ')
    device=next(net.parameters()).device
    yp=predict_model(net,processor.transform(gx),device)
    y_true=gy[ordered_cols(gy,PREFIXES['gain'])].apply(pd.to_numeric,errors='coerce').fillna(0).to_numpy(dtype=np.float32)
    mask=extract_mask(gx)
    baseline=(pd.to_numeric(gx['target_gain'],errors='coerce').fillna(0).to_numpy()[:,None]
              +.5*pd.to_numeric(gx['target_gain_tilt'],errors='coerce').fillna(0).to_numpy()[:,None]
              *np.linspace(-1.,1.,N_CH)[None,:])
    score,bycat=evaluate_predictions(out,'testground',gx,y_true,yp,mask,baseline=baseline)
    (out/'testground_metrics.json').write_text(json.dumps({'ALL':score,**bycat},indent=2,ensure_ascii=False))
    print('POST-CONTEST TestGround metrics (NOT live hidden leaderboard):')
    print(json.dumps({'ALL':score,**bycat},indent=2,ensure_ascii=False))


def run(args):
    root=Path(args.data_root)
    out=Path(args.output);out.mkdir(parents=True,exist_ok=True)
    if args.audit_only:
        ckpt=Path(args.checkpoint) if args.checkpoint else out/'best_physics_edfa.pt'
        if not ckpt.exists():raise FileNotFoundError(f'Cannot audit without saved checkpoint: {ckpt}')
        net,processor=checkpoint_to_model(ckpt,cpu=args.cpu)
        evaluate_testground(root,out,net,processor)
        print(f'Finished without retraining. Results: {out.resolve()}')
        return
    train_x=read_table(root/'Features/Train/train_features.csv')
    train_y=read_table(root/'Features/Train/train_labels.csv')
    print(f'Training rows: {len(train_x):,} | shape: {train_x.shape} | output shape: {train_y.shape}')
    net,processor,metrics,history,vi,yval,pval,mval=train_model(train_x,train_y,args)
    history.to_csv(out/'training_history.csv',index=False)
    plot_diagnostics(out,history,yval,pval,mval)
    # Transparent no-ML reference: gain setpoint + half-span tilt. Same holdout.
    x_coord=np.linspace(-1.,1.,N_CH)[None,:]
    ref_gain=pd.to_numeric(train_x.iloc[vi]['target_gain']).to_numpy()[:,None]
    ref_tilt=pd.to_numeric(train_x.iloc[vi]['target_gain_tilt']).to_numpy()[:,None]
    baseline_metrics=official_score(yval,ref_gain+.5*ref_tilt*x_coord,mval)
    detailed,bycat=evaluate_predictions(out,'validation',train_x.iloc[vi],yval,pval,mval,
                                        baseline=ref_gain+.5*ref_tilt*x_coord)
    (out/'baseline_vs_model.json').write_text(json.dumps(
        {'physics_baseline':baseline_metrics,'learned_physics_model':metrics},indent=2))
    print('Extended validation metrics:',json.dumps(detailed,ensure_ascii=False,indent=2))
    print('Plain Gain/Tilt baseline official score:', round(baseline_metrics['score'],6))
    (out/'validation_by_category.json').write_text(json.dumps(bycat,indent=2))
    print('Validation by category:',json.dumps(bycat,indent=2))
    torch.save({'state_dict':net.state_dict(),
                'input_scaler_mean':torch.as_tensor(processor.scaler.mean_,dtype=torch.float64),
                'input_scaler_scale':torch.as_tensor(processor.scaler.scale_,dtype=torch.float64),
                'config':{'global_dim':len(processor.scaler.mean_), 'hidden':args.hidden}},out/'best_physics_edfa.pt')
    (out/'validation_metrics.json').write_text(json.dumps(metrics,ensure_ascii=False,indent=2))
    device=next(net.parameters()).device
    testfile=root/'Features/Test/test_features.csv'
    if testfile.exists():
        testdf=read_table(testfile)
        testpred=predict_model(net,processor.transform(testdf),device)
        submission=make_submission(testdf,testpred,out/'submission.csv')
        print(f'Submission generated: {out/"submission.csv"}, shape={submission.shape}')
    if args.evaluate_testground:
        evaluate_testground(root,out,net,processor)
    print(f'Files saved to: {out.resolve()}')


def parse_args():
    ap=argparse.ArgumentParser(description='Gain/tilt baseline plus 95-channel Transformer residual')
    ap.add_argument('--data-root',type=str,default='ofc-ml-challenge-data-code')
    ap.add_argument('--output',type=str,default='transformer_results')
    ap.add_argument('--epochs',type=int,default=250)
    ap.add_argument('--patience',type=int,default=40)
    ap.add_argument('--batch-size',type=int,default=64)
    ap.add_argument('--hidden',type=int,default=64)
    ap.add_argument('--lr',type=float,default=0.002)
    ap.add_argument('--val-frac',type=float,default=0.2)
    ap.add_argument('--seed',type=int,default=SEED)
    ap.add_argument('--cpu',action='store_true')
    ap.add_argument('--evaluate-testground',action='store_true',help='Post-contest audit only, never used for fitting')
    ap.add_argument('--audit-only',action='store_true',help='Evaluate saved checkpoint against post-contest TestGround without retraining')
    ap.add_argument('--checkpoint',type=str,default=None,help='Checkpoint file for --audit-only, defaults to output/best_physics_edfa.pt')
    return ap.parse_args()


if __name__=='__main__':run(parse_args())
