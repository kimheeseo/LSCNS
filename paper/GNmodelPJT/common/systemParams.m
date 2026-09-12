function [signal, sys] = systemParams(deltaf, nCh, power_dBm, article)
%SYSTEMPARAMS Executable reconstruction of Appendix-A parameter presets.
if nargin < 1 || isempty(deltaf), deltaf = 0.05; end
if nargin < 2 || isempty(nCh), nCh = 9; end
if nargin < 3 || isempty(power_dBm), power_dBm = 0; end
if nargin < 4, article = 'default'; end
switch lower(string(article))
    case {"default","gn","1"}
        rolloff=0.01; Rs=32e-3; Ls=80; Ns=10; F=5; G=16;
    case {"field_trial","2"}
        rolloff=0.01; Rs=32e-3; Ls=68; Ns=9; F=5.5; G=13.6;
    case {"orange","3"}
        rolloff=0.01; Rs=32e-3; Ls=100; Ns=20; F=5.5; G=20;
    case {"coriant","4"}
        rolloff=0.01; Rs=32e-3; Ls=80; Ns=5; F=5.5; G=16;
    case {"ssmf_15ch","5"}
        rolloff=0.01; Rs=32e-3; Ls=55; Ns=121; F=5.2; G=9.9;
    case {"fiber_ssmf","7"}
        rolloff=0.01; Rs=15.625e-3; Ls=51.06; Ns=1; F=5.5; G=9.7014;
    otherwise
        error('Unknown article preset: %s', article);
end
pW = 1e-3*10.^(power_dBm/10);
if isscalar(pW), pW = repmat(pW,1,nCh); end
BwCh=(1+rolloff)*Rs; Bw=nCh*deltaf;
fCh=(-Bw/2-deltaf/2)+deltaf*(1:nCh);
signal=struct('rolloff',rolloff,'Rs',Rs,'power',pW,'BwCh',BwCh);
sys=struct('Flin',10^(F/10),'Glin',10^(G/10),'Ls',Ls,'nCh',nCh,...
    'deltaf',deltaf,'Bw',Bw,'Ns',Ns,'fCh',fCh,'fo',299792.458/1550);
end
