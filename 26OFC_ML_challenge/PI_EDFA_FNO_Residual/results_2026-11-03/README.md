# PI-EDFA + FNO: official TestGround evaluation, 2026-11-03

**Source:** [GitHub Actions successful run](https://github.com/kimheeseo/LSCNS/actions/runs/41540664667) · [Download complete results artifact](https://github.com/kimheeseo/LSCNS/actions/runs/41540664667/artifacts/13185463032)

## Conditions
- Official OFC 2026 Git LFS CSV measurements
- Recent Train: 397 fit + 93 grouped holdout, seed 2026
- PI-EDFA training: at most 60 epochs; FNO residual training: at most 60 epochs; patience 15
- PI backbone frozen while training FNO correction, COSMOS pretraining **disabled**
- Public TestGround **21,056 samples** used only after checkpoint selection: Aging 4,400; SHB 864; Unseen 15,792.
- Values computed by code implementing the official-style active-channel scoring. **Not a Kaggle hidden-leaderboard submission.**

## Public TestGround scores (lower is better)
| Category | PI-EDFA only | PI-EDFA + FNO | Improvement |
|---|---:|---:|---:|
| AGING | 0.149418 | **0.147972** | 0.97% |
| SHB | 0.149192 | **0.147722** | 0.99% |
| UNSEEN | 0.450383 | **0.449606** | 0.17% |
| **Overall** | 0.393592 | **0.392948** | 0.16% |

## Internal validation
- Recent Train group holdout: **93 rows**; PI-EDFA **0.228532**; PI-EDFA+FNO **0.226595**
- Full categories and detailed MAE/STD/P95/max-error are in the JSON files below.
- Improvements were small; the Unseen result remains roughly **0.45**, indicating substantial room for further generalization work.

## Permanent score files
- [testground_comparison.json](testground_comparison.json): public TestGround, both models and all categories
- [validation_comparison.json](validation_comparison.json): grouped Train holdout scores, sizes, seed, training metadata
- [testground_scores.csv](testground_scores.csv): quick score comparison and relative improvements
- Full model checkpoint (`.pt`), prediction CSV, and training histories are provided by the [Actions artifact](https://github.com/kimheeseo/LSCNS/actions/runs/41540664667/artifacts/13185463032), not committed directly to Git because they are generated/binary outputs.

**Reproducibility:** [GitHub Actions workflow](../../../.github/workflows/ofc-pi-fno-evaluate.yml) · [model.py](../model.py) · [train.py](../train.py). This is one training seed, not a multi-seed confidence estimate.
