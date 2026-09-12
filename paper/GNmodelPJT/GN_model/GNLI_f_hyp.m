function GNLI = GNLI_f_hyp(fiberParams, signal, sys, f, gtx)
%GNLI_F_HYP Hyperbolic-coordinate GN-model entry point.
% Uses the mathematically equivalent Cartesian integral as an executable fallback.
if nargin<5, gtx=[]; end
GNLI=GNLI_f(fiberParams,signal,sys,f,gtx);
end
