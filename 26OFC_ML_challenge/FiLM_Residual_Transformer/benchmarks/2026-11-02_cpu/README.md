# Actual OFC 2026 Residual-FiLM Transformer training and evaluation

GitHub Actions CPU run. COSMOS 12,000 training samples (source subset), 18 max pretraining epochs; latest Train 150 max epochs; seed 2026; 64 hidden, 2 layers, 4 heads; LR 0.002; curvature 0.02. Scratch uses same target split, seed, and LR.

**Caution:** This is NOT full COSMOS training. TestGround is used ONLY after frozen models were selected. Comparing scores across different splits is not fair.

## Latest Train holdout (lower better)

| Model | Overall | Aging | SHB | Unseen |
|---|---:|---:|---:|---:|
| COSMOS -> Train | 4.140808 | 0.395008 | 12.184726 | 0.532990 |
| Train only | 4.137783 | 0.242842 | 12.254903 | 0.640340 |

## Frozen public TestGround (post-selection; lower better)

| Model | Overall | Aging | SHB | Unseen |
|---|---:|---:|---:|---:|
| COSMOS -> Train | 1.572768 | 1.314756 | 1.447053 | 1.608132 |
| Train only | 1.540440 | 1.342096 | 1.539642 | 1.580112 |

## Reproducibility

- Full history: `cosmos_pretrain_history.csv`, `recent_history.csv`, `scratch_history.csv`.
- Same Train split and target LR for transfer and scratch; TestGround labels excluded from training, scaler fit and epoch selection.
- `recent_per_sample.csv` and `scratch_per_sample.csv` list error by held-out Train sample.
- Best checkpoints and `submission.csv` provided as GitHub Actions artifacts, not committed due to file size.
- This is one seed and one curvature weight; further ablations needed before robust claims.

## Observed limitations and immediate diagnosis

- The target Train holdout has two extreme SHB outliers: **ID 219** (max active-channel error **30.5405 dB**), **ID 223** (**30.2297 dB**). Both have `target_gain=18 dB`, 33 active channels. These two observations dominate the internal validation penalty; this is not evidence of a failed download.
- The new `data.py` sends only `pin_dbm`, WSS mask, aging, target gain and target tilt into the network. In contrast, the earlier PI-EDFA/Transformer used the official `EDFA_input_power_total` and **`EDFA_output_power_total`** as global feature inputs. This input regression is **confirmed in code** and is a plausible explanation for the failure to distinguish unusual SHB conditions. Whether it accounts for the observed errors requires an ablation with matched train/test conditions.
- Original scratch/transfer TestGround performances (1.5404 vs 1.5728) are far worse than historical PI-EDFA / Transformer reports. The structural changes should NOT be advertised as accuracy improvements yet.
- Recommended next experiment: preserve residual + per-block FiLM + curvature objective, but incorporate **input/output total powers**, with consistent source/target scaling; test `curvature=0` vs `0.02` and inspect ID 219/223 predictions. Avoid using TestGround for hyperparameter tuning.
