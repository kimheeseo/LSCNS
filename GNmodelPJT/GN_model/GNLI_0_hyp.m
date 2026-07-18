function GNLI0 = GNLI_0_hyp(fiberParams, signal, sys)
%GNLI_0_HYP GN-model NLI PSD at the baseband center f=0 [W/Hz].
GNLI0=GNLI_f_hyp(fiberParams,signal,sys,0,[]);
end
