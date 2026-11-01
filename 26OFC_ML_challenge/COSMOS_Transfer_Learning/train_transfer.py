#!/usr/bin/env python3
"""OFC 2026 COSMOS pretraining + recent Train fine-tuning (PI-EDFA / Transformer).

Keeps both original architectures unchanged. TestGround is used ONLY as an
optional, post-selection audit after all training and model selection.
"""
from __future__ import annotations

import argparse
import copy
import json
import random
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import torch
from sklearn.model_selection import GroupShuffleSplit, train_test_split
from torch.utils.data import DataLoader, TensorDataset

HERE = Path(__file__).resolve().parent
BASE = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(BASE / "PI_EDFA_Transformer"))

import physics_edfa
import transformer_edfa

MODELS = {"pi-edfa": physics_edfa, "transformer": transformer_edfa}
N_CH = 95


def read_pair(root: Path, stem: str):
    source = root / "Features" / "Train"
    feature_path = source / (stem + "_features.csv")
    label_path = source / (stem + "_labels.csv")
    df = transformer_edfa.read_table(feature_path).reset_index(drop=True)
    labels = transformer_edfa.read_table(label_path).reset_index(drop=True)
    if len(df) != len(labels):
        raise ValueError(f"Different row counts for {stem}: {len(df)} vs {len(labels)}")
    if "ID" in df and "ID" in labels and not df.ID.equals(labels.ID):
        raise ValueError(f"{stem}: feature/label IDs do not align")
    # Fail early if this dataset's 95-channel schema is not compatible.
    transformer_edfa.ordered_cols(df, transformer_edfa.PREFIXES["input"])
    transformer_edfa.ordered_cols(df, transformer_edfa.PREFIXES["mask"])
    cols = transformer_edfa.ordered_cols(labels, transformer_edfa.PREFIXES["gain"])
    y = labels[cols].apply(pd.to_numeric, errors="coerce").to_numpy(np.float32)
    if not np.isfinite(y).all():
        # NaNs on inactive channels are tolerated, but not on active labels.
        mask = transformer_edfa.extract_mask(df) > 0.5
        if not np.isfinite(y[mask]).all():
            raise ValueError(f"{stem}: active gain labels contain NaN or Inf")
        y = np.where(mask, y, 0.0).astype(np.float32)
    return df, y


def split_indices(module, df, seed: int, val_frac: float):
    mask = module.extract_mask(df)
    groups = module.get_groups(df, mask)
    index = np.arange(len(df))
    if len(set(groups)) > 1:
        splitter = GroupShuffleSplit(n_splits=1, test_size=val_frac, random_state=seed)
        it, iv = next(splitter.split(index, groups=groups))
    else:
        it, iv = train_test_split(index, test_size=val_frac, random_state=seed)
    if not len(it) or not len(iv):
        raise ValueError("Empty training or validation split")
    return it, iv


def subset_indices(indices, max_rows: int, seed: int):
    if max_rows <= 0 or max_rows >= len(indices):
        return indices
    return np.sort(np.random.default_rng(seed).choice(indices, max_rows, replace=False))


def create_loader(features, y, indices, batch_size, shuffle):
    arrays = [*features, y, features[2]]  # inputs (six) + labels + mask
    # Tensors are kept on CPU and transferred batchwise to avoid GPU OOM.
    tensors = [torch.as_tensor(np.ascontiguousarray(a[indices]), dtype=torch.float32)
               for a in arrays]
    return DataLoader(TensorDataset(*tensors), batch_size=batch_size,
                      shuffle=shuffle, num_workers=0, pin_memory=False)


@torch.no_grad()
def evaluate_loader(net, loader, module, device):
    net.eval()
    preds, actual, masks = [], [], []
    for batch in loader:
        inp = [b.to(device) for b in batch[:6]]
        p, _ = net(*inp)
        preds.append(p.detach().cpu().numpy())
        actual.append(batch[6].numpy())
        masks.append(batch[7].numpy())
    return module.official_score(np.concatenate(actual), np.concatenate(preds),
                                 np.concatenate(masks))


def train_stage(net, module, processor, x, y, train_idx, val_idx, args, stage, out,
                epochs, lr, patience, batch_size, device):
    transformed = processor.transform(x)
    tr = create_loader(transformed, y, train_idx, batch_size, True)
    va = create_loader(transformed, y, val_idx, max(batch_size, 256), False)
    optimizer = torch.optim.AdamW(net.parameters(), lr=lr, weight_decay=2e-4)
    scheduler = torch.optim.lr_scheduler.ReduceLROnPlateau(
        optimizer, patience=8, factor=0.6, mode="min")
    best, bad, best_state, history = float("inf"), 0, None, []
    for epoch in range(1, epochs + 1):
        net.train()
        train_losses = []
        for batch in tr:
            b = [a.to(device) for a in batch]
            optimizer.zero_grad(set_to_none=True)
            pred, components = net(*b[:6])
            loss = module.loss_fn(pred, b[6], b[7], components)
            if not torch.isfinite(loss):
                raise RuntimeError(f"Non-finite loss in {stage} epoch {epoch}")
            loss.backward()
            torch.nn.utils.clip_grad_norm_(net.parameters(), 3.)
            optimizer.step()
            train_losses.append(loss.item())
        metrics = evaluate_loader(net, va, module, device)
        score = metrics["score"]
        scheduler.step(score)
        history.append({"stage": stage, "epoch": epoch,
                        "train_loss": float(np.mean(train_losses)),
                        "val_score": score, "val_mae": metrics["mae"],
                        "learning_rate": optimizer.param_groups[0]["lr"]})
        if score < best - 1e-5:
            best = score
            bad = 0
            best_state = copy.deepcopy({k: v.detach().cpu() for k, v in net.state_dict().items()})
        else:
            bad += 1
        if epoch == 1 or epoch % 10 == 0 or epoch == epochs:
            print(f"{stage} {epoch:3d}/{epochs}: val={score:.6f} best={best:.6f}",
                  flush=True)
        if bad >= patience:
            print(f"{stage}: early stop at epoch {epoch}", flush=True)
            break
    net.load_state_dict(best_state)
    # Report metrics of the BEST selected checkpoint, not the last training epoch.
    selected_metrics = evaluate_loader(net, va, module, device)
    ckpt = {"model": args.model, "stage": stage, "state_dict": best_state,
            "input_scaler_mean": torch.tensor(processor.scaler.mean_, dtype=torch.float64),
            "input_scaler_scale": torch.tensor(processor.scaler.scale_, dtype=torch.float64),
            "config": {"global_dim": int(len(processor.scaler.mean_)),
                       "hidden": int(args.hidden)}, "seed": args.seed}
    torch.save(ckpt, out / (stage + ".pt"))
    pd.DataFrame(history).to_csv(out / (stage + "_history.csv"), index=False)
    return selected_metrics, history


@torch.no_grad()
def predict(net, processor, module, df, device):
    return module.predict_model(net, processor.transform(df), device=device)


def category_score(module, df, y, predictions):
    mask = module.extract_mask(df)
    summary = {"ALL": module.official_score(y, predictions, mask)}
    cat_col = next((c for c in df.columns if c.lower() == "category"), None)
    if cat_col:
        categories = df[cat_col].astype(str).str.lower().to_numpy()
        for cat in sorted(set(categories)):
            on = categories == cat
            summary[cat] = module.official_score(y[on], predictions[on], mask[on])
    return summary


def evaluate_final(args, net, processor, module, root, out, name, device):
    ground = root / "Features" / "TestGround"
    gx = transformer_edfa.read_table(ground / "test_features.csv").reset_index(drop=True)
    gy = transformer_edfa.read_table(ground / "test_labels.csv").reset_index(drop=True)
    if len(gx) != len(gy):
        raise ValueError("TestGround feature and label row counts differ")
    if "ID" in gx and "ID" in gy and not gx.ID.equals(gy.ID):
        raise ValueError("TestGround feature/label IDs do not align")
    ycols = module.ordered_cols(gy, module.PREFIXES["gain"])
    actual = gy[ycols].apply(pd.to_numeric, errors="coerce").to_numpy(np.float32)
    ground_mask = module.extract_mask(gx).astype(bool)
    if not np.isfinite(actual[ground_mask]).all():
        raise ValueError("Non-finite active TestGround label")
    actual = np.where(ground_mask, actual, 0.0)
    result = category_score(module, gx, actual, predict(net, processor, module, gx, device))
    (out / (name + "_testground.json")).write_text(json.dumps(result, indent=2))
    return result


def model_run(args, model_name, root, recent_x, recent_y, cosmos_x, cosmos_y,
              recent_train_idx, recent_val_idx, cosmos_train_idx, cosmos_val_idx):
    module = MODELS[model_name]
    # Small data and large data must share the same feature convention.
    module.ordered_cols(cosmos_x, module.PREFIXES["input"])
    module.ordered_cols(recent_x, module.PREFIXES["input"])
    module.seed_everything(args.seed)
    device = "cpu" if args.cpu or not torch.cuda.is_available() else "cuda"
    out = Path(args.output) / model_name
    out.mkdir(parents=True, exist_ok=True)
    hidden = args.hidden if args.hidden else (96 if model_name == "pi-edfa" else 64)
    local_args = copy.copy(args)
    local_args.model = model_name
    local_args.hidden = hidden

    def fresh_model(processor):
        return module.PhysicalSpectralNet(
            global_dim=len(processor.scaler.mean_), hidden=hidden).to(device)

    # IMPORTANT: COSMOS scaler fits source train only; it is reused unchanged
    # for fine-tuning so the pretrained network sees consistent feature units.
    source_processor = module.EDFAFeatures().fit(cosmos_x.iloc[cosmos_train_idx])
    net = fresh_model(source_processor)
    source_metrics, source_history = train_stage(
        net, module, source_processor, cosmos_x, cosmos_y, cosmos_train_idx,
        cosmos_val_idx, local_args, "cosmos_pretrain", out, args.pretrain_epochs,
        args.pretrain_lr, args.pretrain_patience, args.pretrain_batch_size, device)

    # Fine-tune ALL layers with a smaller LR. The recent holdout never
    # contributes to optimizer, scaler fitting, or pretraining.
    ft_metrics, ft_history = train_stage(
        net, module, source_processor, recent_x, recent_y, recent_train_idx,
        recent_val_idx, local_args, "recent_finetune", out, args.finetune_epochs,
        args.finetune_lr, args.finetune_patience, args.batch_size, device)
    transfer_pred = predict(net, source_processor, module,
                            recent_x.iloc[recent_val_idx], device)
    transfer_bycat = category_score(module, recent_x.iloc[recent_val_idx],
                                     recent_y[recent_val_idx], transfer_pred)

    results = {"source_val": source_metrics, "transfer_val": ft_metrics,
               "transfer_val_by_category": transfer_bycat,
               "source_rows": int(len(cosmos_train_idx)),
               "recent_train_rows": int(len(recent_train_idx)),
               "recent_val_rows": int(len(recent_val_idx))}
    if args.evaluate_testground:
        results["transfer_testground"] = evaluate_final(
            local_args, net, source_processor, module, root, out, "transfer", device)

    # Fair scratch control: same train/validation indices and seed, but scale
    # using recent training data ONLY and do not load COSMOS weights.
    if not args.skip_scratch:
        module.seed_everything(args.seed)
        scratch_proc = module.EDFAFeatures().fit(recent_x.iloc[recent_train_idx])
        scratch_net = fresh_model(scratch_proc)
        scratch_metric, _ = train_stage(
            scratch_net, module, scratch_proc, recent_x, recent_y,
            recent_train_idx, recent_val_idx, local_args, "recent_scratch",
            out, args.finetune_epochs, args.scratch_lr, args.finetune_patience,
            args.batch_size, device)
        results["scratch_val"] = scratch_metric
        results["scratch_val_by_category"] = category_score(
            module, recent_x.iloc[recent_val_idx], recent_y[recent_val_idx],
            predict(scratch_net, scratch_proc, module,
                    recent_x.iloc[recent_val_idx], device))
        if args.evaluate_testground:
            results["scratch_testground"] = evaluate_final(
                local_args, scratch_net, scratch_proc, module, root, out, "scratch", device)

    (out / "results.json").write_text(json.dumps(results, indent=2, ensure_ascii=False))
    testfile = root / "Features" / "Test" / "test_features.csv"
    if testfile.exists():
        test_x = transformer_edfa.read_table(testfile)
        pred = predict(net, source_processor, module, test_x, device)
        module.make_submission(test_x, pred, out / "submission.csv")
    print(f"FINISHED {model_name}:", json.dumps(results, indent=2)[:2500], flush=True)
    return results


def get_args():
    parser = argparse.ArgumentParser(
        description="COSMOS pretrain -> recent Train fine-tune for both OFC models")
    parser.add_argument("--data-root", default="/content/ofc-ml-challenge-data-code")
    parser.add_argument("--output", default="/content/ofc_cosmos_transfer_results")
    parser.add_argument("--models", nargs="+", default=["pi-edfa", "transformer"],
                        choices=list(MODELS))
    parser.add_argument("--pretrain-epochs", type=int, default=60)
    parser.add_argument("--pretrain-patience", type=int, default=12)
    parser.add_argument("--finetune-epochs", type=int, default=180)
    parser.add_argument("--finetune-patience", type=int, default=30)
    parser.add_argument("--pretrain-lr", type=float, default=0.001)
    parser.add_argument("--finetune-lr", type=float, default=0.0001)
    parser.add_argument("--scratch-lr", type=float, default=0.002)
    parser.add_argument("--pretrain-batch-size", type=int, default=128)
    parser.add_argument("--batch-size", type=int, default=64)
    parser.add_argument("--pretrain-val-frac", type=float, default=0.08)
    parser.add_argument("--recent-val-frac", type=float, default=0.2)
    parser.add_argument("--cosmos-max-rows", type=int, default=0,
                        help="0 = all COSMOS training rows; >0 = deterministic subset")
    parser.add_argument("--cosmos-val-max-rows", type=int, default=3000)
    parser.add_argument("--hidden", type=int, default=0,
                        help="0=original per-model size (PI 96, Transformer 64)")
    parser.add_argument("--seed", type=int, default=2026)
    parser.add_argument("--cpu", action="store_true")
    parser.add_argument("--skip-scratch", action="store_true")
    parser.add_argument("--evaluate-testground", action="store_true",
                        help="Post-selection audit ONLY. Never used for training or tuning")
    args = parser.parse_args()
    if args.pretrain_epochs < 1 or args.finetune_epochs < 1:
        parser.error("Epochs must be positive")
    if not (0 < args.recent_val_frac < 1 and 0 < args.pretrain_val_frac < 1):
        parser.error("Validation fractions must be in (0, 1)")
    return args


def main():
    args = get_args()
    root = Path(args.data_root)
    out = Path(args.output)
    out.mkdir(parents=True, exist_ok=True)
    transformer_edfa.seed_everything(args.seed)
    cosmos_x, cosmos_y = read_pair(root, "COSMOS")
    recent_x, recent_y = read_pair(root, "train")
    print(f"COSMOS rows={len(cosmos_x)}, recent Train rows={len(recent_x)}", flush=True)
    ct, cv = split_indices(transformer_edfa, cosmos_x, args.seed, args.pretrain_val_frac)
    rt, rv = split_indices(transformer_edfa, recent_x, args.seed, args.recent_val_frac)
    ct = subset_indices(ct, args.cosmos_max_rows, args.seed)
    cv = subset_indices(cv, args.cosmos_val_max_rows, args.seed + 1)
    scores = {}
    for name in args.models:
        scores[name] = model_run(
            args, name, root, recent_x, recent_y, cosmos_x, cosmos_y, rt, rv, ct, cv)
    (out / "comparison.json").write_text(
        json.dumps(scores, indent=2, ensure_ascii=False))
    rows = []
    for name, item in scores.items():
        for setting in ("source_val", "transfer_val", "scratch_val"):
            if setting in item:
                rows.append({"model": name, "stage": setting, **item[setting]})
    pd.DataFrame(rows).to_csv(out / "comparison.csv", index=False)
    print("Saved:", out.resolve(), flush=True)


if __name__ == "__main__":
    main()
