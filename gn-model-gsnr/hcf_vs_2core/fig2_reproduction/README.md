# Fig. 2 — Throughput versus launch power

This folder reproduces Fig. 2 of:

> R. Sohanpal, E. Sillekens, M. Jarmolovicius, R. I. Killey, and P. Bayvel,
> “On the Optimum Energy-per-bit Launch Power in Coherent Hollow-core Fibre
> Transmission Systems,” 2026, [arXiv:2606.17942](https://arxiv.org/abs/2606.17942).

The outputs intentionally distinguish an **exact source-data redraw** from an
**independent parameter-based calculation**. Replotting the authors' data
reproduces a figure; matching it from disclosed physical inputs validates the
model.

## Run

From the `gn-model-gsnr` directory:

```bash
python hcf_vs_2core/fig2_reproduction/reproduce_fig2.py
```

Dependencies are Python 3, NumPy, and Matplotlib.

## Reproduced paper configuration

| Quantity | Value |
|---|---:|
| Modulation assumption | Gaussian, dual polarisation |
| Symbol rate | 140 GBd |
| Channel spacing | 150 GHz |
| C-band channels | 29 |
| Transceiver SNR | 20 dB |
| Amplifier noise figure | 5 dB |
| HCF link | 1 x 200 km |
| HCF attenuation at 1550 nm | 0.075 dB/km |
| HCF dispersion at 1550 nm | 3.2 ps/(nm km) |
| HCF nonlinearity | 5e-4 (W km)^-1 |
| HCF inter-modal interference | -52 dB/km |
| SMF link | 3 x 67 km (200 km total) |

For each channel, the script evaluates the paper's additive form

\[
\frac{1}{\mathrm{SNR}}=\frac{P_{\mathrm{ASE}}}{P}
 + \eta P^2 + \frac{P_{\mathrm{IMI}}}{P}
 + \frac{P_{\mathrm{TRN}}}{P},
\]

with the fixed transceiver and HCF-IMI contributions represented directly as
inverse-SNR terms. Total C-band throughput is

\[
T_{\mathrm{tot}}=29\,(2R)\log_2(1+\mathrm{SNR}).
\]

## Results

| Case | Paper source data | Physical-input model | Visible-range RMSE |
|---|---:|---:|---:|
| HCF | 52.670866 Tb/s at 22 dBm/ch | 52.671814 Tb/s at 22.4 dBm/ch | 0.258832 Tb/s |
| SMF | 37.670862 Tb/s at 10 dBm/ch | 52.812451 Tb/s at 3.0 dBm/ch | 20.017368 Tb/s |

The HCF peak throughput is reproduced to within 0.001 Tb/s. Its full visible
curve has an RMSE of 0.259 Tb/s, about 0.49% of the reference peak. The 0.4 dB
peak-power difference is mainly the difference between the paper data's 1 dB
sampling and this script's 0.1 dB grid; the paper source itself also mentions a
22.4 dBm/channel C-band optimum in a commented table row.

The SMF physical-input curve is **not an independent reproduction**. The paper
states the topology and common system settings but does not tabulate the flat
C-band SMF attenuation, dispersion, nonlinearity, or exact wavelength profile
used for Fig. 2. The script therefore exposes and uses the parent repository's
low-loss SMF baseline (`0.15 dB/km`, `-21 ps/(nm km)`, and
`0.81 (W km)^-1`). Those assumptions do not recreate the paper's SMF curve.

As a diagnostic, the script also fits the two free coefficients of the paper's
Eq. (1), `P_ASE` and `eta`, while keeping transceiver SNR and HCF IMI fixed.
This equivalent fit yields visible-range RMSE values of 0.000485 Tb/s for HCF
and 0.004675 Tb/s for SMF. It confirms that the source-data curves follow the
stated ASE-plus-cubic-NLI algorithmic form, but the SMF fit must not be cited as
an independent prediction because it is calibrated to Fig. 2.

## Files

- `reproduce_fig2.py` — self-contained calculation, fitting, CSV, and plotting
- `reference_fig2_hcf.tsv` — exact HCF data from the paper's arXiv source
- `reference_fig2_smf.tsv` — exact SMF data from the paper's arXiv source
- `fig2_reproduction.svg` — exact source-data redraw
- `fig2_model_validation.svg` — source data versus physical model and Eq. (1) fit
- `fig2_reproduction.csv` — point-by-point values and residuals
- `fig2_metrics.csv` — peak values, fitted coefficients, and RMSE metrics

The reference TSV files retain the authors' original `lpch` and `capc` columns.
They were extracted from `Data/Fig2ThroughputC_HCF.txt` and
`Data/Fig2ThroughputC_SMF.txt` in the arXiv source package.
