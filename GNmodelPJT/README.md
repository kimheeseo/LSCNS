# GN Model Project (MATLAB Appendix Transcription)

This folder contains MATLAB code transcribed from **Appendix A - Matlab Files** of the uploaded thesis *Extended GN Model for L-C Band with final experimental verifications* by Angel Alonso Gomez.

## Files

- `GNLI_f_hyp.m` - GN model using hyperbolic coordinates at an arbitrary analysis frequency.
- `GNLI_0_hyp.m` - GN model using hyperbolic coordinates at `f = 0`.
- `GNLI_f.m` - GN model using Cartesian coordinates.
- `GGNLI_f.m` - generalized GN model, matrix approach.
- `GGNLI_f2.m` - generalized GN model, loop/matrix approach.
- `GGNLI_f3.m` - generalized GN model, hybrid approach.
- `systemParams.m` - signal and system parameter presets used in the thesis.
- `fiber_var.m` - Raman/SRS power-profile evolution along a span.
- `fiber.m` - Raman/SRS power profile at the end of a span.

## Important notes

The PDF typesets source code rather than embedding original `.m` files. Therefore, the code was transcribed and normalized from the printed appendix:

- spaced identifiers such as `f u n c t i o n` were restored;
- mathematical glyphs were converted to MATLAB operators;
- wrapped lines were joined;
- function names were aligned with file names.

The appendix does not include every dependency referenced by these functions (for example `GTX` and externally generated fiber profiles). The files should therefore be treated as a research transcription that may still require numerical validation and minor corrections before production use.

## Source

Appendix A, thesis pages 67-97. Original author comments and parameter presets are retained where recoverable from the PDF.
