# COSMOS Transfer PILOT (not full-data result)

Automated GitHub Actions CPU experiment. 4,000 COSMOS train rows sampled; 6 pretraining epochs; 55 maximum recent Train fine-tuning epochs. Seed=2026, same Train holdout for both models and scratch baseline.

| Model | COSMOS pretrain validation | Transfer recent-val | Scratch recent-val | Transfer-vs-scratch |
|---|---:|---:|---:|---:|
| pi-edfa | 0.194916 | 3.622148 | 0.234593 | -1444.0% |
| transformer | 0.207240 | 3.749114 | 0.334849 | -1019.6% |

The above are internal validation scores, not Kaggle final scores. TestGround is NOT evaluated in pilot. Full COSMOS pretraining has NOT been completed.

## Erratum — 2026-11-02 (original values preserved)

**Reporting error discovered and fixed:** the original summary and `comparison.json` stage-level `*_val` values used **last-epoch** scores even though prediction/Category reports used the restored **best checkpoint**. The correct best-checkpoint internal validation comparison for this same original pilot is:

| Model | Transfer best | Scratch best |
|---|---:|---:|
| PI-EDFA | 3.622148 | 0.226496 |
| Transformer | 3.733105 | 0.293547 |

Do not use the original -1444% / -1019.6% as scientific effect estimates: the learning rates were **not controlled** (Transfer 0.0001 vs Scratch 0.002), and the stage summary used inconsistent epochs. See the [notebook investigation](../../COSMOS_Transfer_Error_Analysis_Report.ipynb) and [equal-LR controlled results](../lr_control/README.md). Both experiments use recent Train heldout validation, NOT TestGround.
