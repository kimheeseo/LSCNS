# HCF versus two-core MCF

This folder compares a single-core hollow-core fibre (HCF) link with a two-core
multicore fibre (MCF) link by reusing the parent project's GN/GSNR and submarine
power-feed model.

## Run

From the `gn-model-gsnr` project root:

```bash
python -m pip install numpy matplotlib
python hcf_vs_2core/compare_hcf_vs_2core.py
```

The script regenerates `results.csv` and `capacity_vs_span.svg`.

## Scenario

Both cases use a 6000 km, 12-fibre-pair, C-band link with 100 GBd channels,
112.5 GHz spacing, 4.3 THz occupied bandwidth, an 18 kV PFE, 2.5% amplifier
power-conversion efficiency, a 20 dB transceiver-SNR ceiling, and an 18 dBm
total-WDM launch-power cap per amplifier.

| Parameter | HCF (1 core) | MCF (2 core) |
|---|---:|---:|
| Attenuation | 0.075 dB/km | 0.15 dB/km |
| Dispersion magnitude | 3.2 ps/(nm km) | 21 ps/(nm km) |
| Nonlinearity, gamma | 5e-4 1/(W km) | 0.81 1/(W km) |
| Distributed IMI/XT | -52 dB/km | -80 dB/km |
| FIFO loss | 0 dB/span end | 0.3 dB/span end |
| Amplifier noise figure | 5.0 dB | 4.5 dB |
| Spatial cores per fibre | 1 | 2 |

The HCF values are the 1550 nm C-band assumptions reported by Sohanpal et al.
(2026): 0.075 dB/km attenuation, 3.2 ps/(nm km) dispersion, gamma of
5e-4 1/(W km), and -52 dB/km IMI. The MCF values retain the baseline values used
by the parent repository.

## Reproduced result

| Comparison point | HCF | Two-core MCF |
|---|---:|---:|
| Peak within 50-250 km scan | 196.916 Tb/s at 50 km | 416.554 Tb/s at 50 km |
| System GSNR at 50 km | 12.78 dB | 13.56 dB |
| Capacity at 170 km | 184.886 Tb/s | 178.447 Tb/s |
| Capacity at 200 km | 177.517 Tb/s | 83.839 Tb/s |
| Repeaters at 200 km | 29 | 29 |

The two-core MCF has the higher unconstrained peak capacity because it provides
twice the spatial channels. HCF loses that short-span comparison despite its lower
loss and nonlinearity. As span length increases, however, MCF ASE and nonlinear
penalties rise rapidly. HCF becomes higher-capacity at approximately 170 km and
retains useful GSNR into the 200-250 km range.

## Interpretation limits

- HCF IMI is mapped to the existing model's distributed `gamma_XT` term; this is
  an approximation, not a full modal-noise model.
- The 20 dB transceiver SNR is combined after optical-link optimization because
  the parent optimizer has no transceiver-noise input.
- The absolute Tb/s values follow the parent repository's Shannon-like capacity
  convention. The relative trends are more reliable than the absolute capacity.
- Wavelength-dependent HCF loss, gas absorption, splice MPI, gain tilt, ISRS, and
  modulation/FEC thresholds are not included.

## HCF parameter source

R. Sohanpal et al., *On the Optimum Energy-per-bit Launch Power in Coherent
Hollow-core Fibre Transmission Systems*, arXiv:2606.17942, 2026.
