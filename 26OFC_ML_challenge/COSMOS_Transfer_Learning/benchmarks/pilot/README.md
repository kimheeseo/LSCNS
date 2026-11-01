# COSMOS Transfer PILOT (not full-data result)

Automated GitHub Actions CPU experiment. 4,000 COSMOS train rows sampled; 6 pretraining epochs; 55 maximum recent Train fine-tuning epochs. Seed=2026, same Train holdout for both models and scratch baseline.

| Model | COSMOS pretrain validation | Transfer recent-val | Scratch recent-val | Transfer-vs-scratch |
|---|---:|---:|---:|---:|
| pi-edfa | 0.194916 | 3.622148 | 0.234593 | -1444.0% |
| transformer | 0.207240 | 3.749114 | 0.334849 | -1019.6% |

The above are internal validation scores, not Kaggle final scores. TestGround is NOT evaluated in pilot. Full COSMOS pretraining has NOT been completed.
