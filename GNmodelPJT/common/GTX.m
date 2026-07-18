function gtx = GTX(signal, sys, f)
%GTX Rectangular WDM power spectral density.
% f is in THz; output is W/THz.
validateattributes(f, {'numeric'}, {'vector','real','finite'});
gtx = zeros(size(f));
for k = 1:sys.nCh
    halfBW = signal.BwCh/2;
    inband = abs(f - sys.fCh(k)) <= halfBW;
    gtx(inband) = gtx(inband) + signal.power(k)/signal.BwCh;
end
end
