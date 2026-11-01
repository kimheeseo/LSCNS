# Equal Learning Rate COSMOS Transfer Diagnostic

Reproduces pilot using 4,000 COSMOS samples / 6 epochs pretrain / 55 epochs fine-tuning / seed 2026 / same recent Train split. Both transfer and scratch use LR=0.002 (pilot transfer LR=0.0001).
All score values in this report use **best restored checkpoint** metrics, unlike the earlier pilot summary bug.

| Model | Transfer best recent-val | Scratch best recent-val | Difference (transfer-scratch) |
|---|---:|---:|---:|
| pi-edfa | 0.219143 | 0.231684 | -0.012541 |
| transformer | 0.302495 | 0.274992 | +0.027503 |

See each model folder for per-row SHB error, worst channels, and 10 COSMOS-scaled feature z-score summaries. These are internal Train validation, not public TestGround.
