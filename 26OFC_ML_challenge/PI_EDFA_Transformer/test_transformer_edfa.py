#!/usr/bin/env python3
"""Full-COSMOS train/validation/test evaluation for transformer_edfa.py.

The official COSMOS raw dataset is not the competition's Aging/SHB/Unseen
TestGround. This script trains on all available COSMOS samples after a
reproducible train/validation/test split and reports the official scoring
formula overall and by category when category labels exist. Unlabelled COSMOS
rows are honestly reported as category='COSMOS (unlabelled)'.
"""
from __future__ import annotations
import argparse, json, sys, time
from pathlib import Path
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import transformer_edfa as model


def locate_csv(root: Path, name: str) -> Path:
    candidates = list(root.rglob(name))
    if not candidates:
        raise FileNotFoundError(f"Could not find {name} under {root}")
    return candidates[0]


class ProgressReporter:
    """Report approximate end-to-end milestones every 10%.

    0–10%: CSV preparation and splitting
    10–90%: up to the configured maximum training epochs
    90–100%: holdout evaluation and artifact export
    These are work-stage milestones, not measured elapsed-time percentages.
    """
    def __init__(self):
        self.started = time.monotonic()
        self.next_mark = 10

    def reach(self, percent, description):
        while self.next_mark <= min(100, percent):
            elapsed = int(time.monotonic() - self.started)
            h, rest = divmod(elapsed, 3600)
            m, sec = divmod(rest, 60)
            print(f"[진행률 {self.next_mark:3d}% / 100%] {description} | 경과 {h:02d}:{m:02d}:{sec:02d}", flush=True)
            self.next_mark += 10


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-root", required=True,
                    help="Root of the official COSMOS-EDFA dataset or OFC data repo")
    ap.add_argument("--features", default=None, help="Explicit COSMOS feature CSV")
    ap.add_argument("--labels", default=None, help="Explicit COSMOS label CSV")
    ap.add_argument("--output", default=str(HERE / "cosmos_full_test_results"))
    ap.add_argument("--epochs", type=int, default=250)
    ap.add_argument("--patience", type=int, default=40)
    ap.add_argument("--test-frac", type=float, default=.15)
    ap.add_argument("--val-frac", type=float, default=.15,
                    help="Fraction of remaining train+validation rows")
    ap.add_argument("--seed", type=int, default=2026)
    ap.add_argument("--cpu", action="store_true")
    args = ap.parse_args()

    progress = ProgressReporter()
    print("[시작] COSMOS CSV 불러오기 및 데이터 분할 중... (전체 진행률은 단계/epoch 기준입니다.)", flush=True)
    root = Path(args.data_root)
    fx = Path(args.features) if args.features else locate_csv(root, "COSMOS_features.csv")
    fy = Path(args.labels) if args.labels else locate_csv(root, "COSMOS_labels.csv")
    x = model.read_table(fx)
    ydf = model.read_table(fy)
    if len(x) != len(ydf):
        raise ValueError(f"Feature/label row mismatch: {len(x)} vs {len(ydf)}")
    if len(x) != 202753:
        print(f"NOTE: loaded {len(x):,} rows (expected approximately 202,753); continuing.")

    indices = np.arange(len(x))
    trainval_i, test_i = train_test_split(indices, test_size=args.test_frac,
                                          random_state=args.seed, shuffle=True)
    train_i, val_i = train_test_split(trainval_i, test_size=args.val_frac,
                                      random_state=args.seed, shuffle=True)
    train_x, train_y = x.iloc[train_i].reset_index(drop=True), ydf.iloc[train_i].reset_index(drop=True)
    val_x, val_y = x.iloc[val_i].reset_index(drop=True), ydf.iloc[val_i].reset_index(drop=True)
    test_x, test_y = x.iloc[test_i].reset_index(drop=True), ydf.iloc[test_i].reset_index(drop=True)

    progress.reach(10, f"CSV 로드 및 분할 완료 (총 {len(x):,}행)")
    print(f"[학습 준비] Transformer 학습 시작: 최대 {args.epochs} Epoch (조기 종료 가능).", flush=True)

    # Reuse the model's own training path; its internal validation split is
    # replaced with a temporary train-only frame, while the external validation
    # and test sets remain untouched for final reporting.
    out = Path(args.output); out.mkdir(parents=True, exist_ok=True)
    cli = argparse.Namespace(seed=args.seed, cpu=args.cpu, val_frac=args.val_frac,
        batch_size=64, hidden=64, lr=.002, epochs=args.epochs, patience=args.patience)
    def on_epoch_end(epoch, planned_epochs, val_score, best_score):
        percent = 10 + int(80 * epoch / max(1, planned_epochs))
        progress.reach(percent,
                       f"Transformer 학습 Epoch {epoch}/{planned_epochs}, 검증점수 {val_score:.5f}, 최고 {best_score:.5f}")

    net, processor, _, history, *_ = model.train_model(
        train_x, train_y, cli, progress_callback=on_epoch_end
    )
    trained_epochs = len(history)
    train_summary = (f"Transformer 학습 완료 ({trained_epochs}/{args.epochs} Epoch)"
                     if trained_epochs >= args.epochs
                     else f"Early Stopping: {trained_epochs}/{args.epochs} Epoch에서 학습 종료")
    progress.reach(90, train_summary)
    print("[평가] Validation/Test 예측 및 점수·그래프 저장 중...", flush=True)

    device = next(net.parameters()).device
    result = {"rows": {"total": len(x), "train": len(train_x), "validation": len(val_x), "test": len(test_x)},
              "split": {"seed": args.seed, "test_fraction": args.test_frac,
                        "validation_fraction_of_trainval": args.val_frac}}
    for name, xx, yy in (("validation", val_x, val_y), ("test", test_x, test_y)):
        mask = model.extract_mask(xx)
        pred = model.predict_model(net, processor.transform(xx), device)
        truth = yy[model.ordered_cols(yy, model.PREFIXES["gain"])].apply(
            pd.to_numeric, errors="coerce").fillna(0).to_numpy(dtype=np.float32)
        truth = np.where(mask > .5, truth, 0.)
        metadata = xx.copy()
        if not any(c.lower() == "category" for c in metadata.columns):
            metadata["Category"] = "COSMOS (unlabelled)"
        overall, bycat = model.evaluate_predictions(out, name, metadata, truth, pred, mask)
        result[name] = {"overall": overall, "by_category": bycat}
        print(f"\n{name.upper()} OFC-style score")
        print(json.dumps(result[name], indent=2, ensure_ascii=False))
    history.to_csv(out / "training_history.csv", index=False)
    (out / "cosmos_split_scores.json").write_text(json.dumps(result, indent=2, ensure_ascii=False))
    progress.reach(100, f"Validation/Test 평가·결과 저장 완료: {out.resolve()}")
    print(f"\\nSaved results to {out.resolve()}", flush=True)
    print("Note: COSMOS does not carry the OFC Aging/SHB/Unseen benchmark labels; those category scores require the labelled official TestGround.")


if __name__ == "__main__":
    main()
