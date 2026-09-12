clear; clc;
root=fileparts(fileparts(mfilename('fullpath')));
addpath(fullfile(root,'common'),fullfile(root,'GN_model'),fullfile(root,'GGN_model'));
[signal,sys]=systemParams(0.05,9,0,'default');
fp=fiberParams('SSMF');
n=1201; freq=linspace(-sys.Bw/2,sys.Bw/2,n); psd=GTX(signal,sys,freq);
gn=GNLI_f(fp,signal,sys,0,psd);
zN=101; J=sum(signal.power); Cr=0.028; fMax=15;
fibProf=zeros(sys.nCh,zN); alpha=repmat(fp.alpha,1,sys.nCh);
for k=1:sys.nCh
    fibProf(k,:)=fiber_var(signal.power,Cr,fMax,alpha,sys.deltaf,sys.Ls,k,sys.nCh,J,zN);
end
ggn=GGNLI_f_2(fp,signal,sys,0,fibProf,psd,zN);
fprintf('GN NLI PSD  : %.4e W/Hz\n',gn);
fprintf('GGN NLI PSD : %.4e W/Hz\n',ggn);
