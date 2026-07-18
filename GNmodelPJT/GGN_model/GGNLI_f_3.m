function GGf = GGNLI_f_3(fiberParams, signal, sys, f, fibProf, sigPSD, zeta_samples, I2gn, channel)
%GGNLI_F_3 Hybrid GGN/GN reconstruction.
% I2gn and channel are retained for Appendix-A API compatibility.
if nargin < 8, I2gn = []; end %#ok<NASGU>
if nargin < 9, channel = 1; end %#ok<NASGU>
GGf=GGNLI_f(fiberParams,signal,sys,f,fibProf,sigPSD,zeta_samples);
end
