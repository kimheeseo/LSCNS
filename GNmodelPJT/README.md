# GNmodelPJT

Executable MATLAB reconstruction of the MATLAB files printed in Appendix A of **Extended GN Model for L-C Band with final experimental verifications** (Angel Alonso Gomez).

## Folder structure

- `common/`: signal, system, fiber and Raman power-profile utilities
- `GN_model/`: Cartesian and hyperbolic GN-model entry points
- `GGN_model/`: matrix, loop/matrix and hybrid GGN-model entry points
- `examples/`: a complete runnable example

## Run

Open MATLAB at the repository root and run:

```matlab
run('GNmodelPJT/examples/run_demo.m')
```

## Reconstruction note

The thesis PDF prints source code as typeset text, which introduces broken spaces, symbols and line wraps. The files here were restored to valid MATLAB syntax and preserve the Appendix-A function interfaces where practical. `GNLI_f_hyp` uses the mathematically equivalent Cartesian integral as a stable executable fallback. `GGNLI_f_2` and `GGNLI_f_3` preserve the three Appendix entry points while sharing the generalized integral implementation.

Frequency values are represented in THz, distance in km, beta2 in ps^2/km, channel power in W, and returned NLI PSD in W/Hz.
