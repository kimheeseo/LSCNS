function GNLI = GNLI_f(fiberParams, signal, sys, f, gtx)
%GNLI_F Cartesian-coordinate GN-model NLI PSD at frequency f [W/Hz].
if nargin<5 || isempty(gtx)
    nSamples=max(801,floor(sys.Bw*23000/8));
    grid=linspace(-sys.Bw/2,sys.Bw/2,nSamples); gtx=GTX(signal,sys,grid);
else
    nSamples=numel(gtx); grid=linspace(-sys.Bw/2,sys.Bw/2,nSamples);
end
f1=grid; f2=grid; I2=zeros(size(f1));
[~,index]=min(abs(f-f2)); da=-2*fiberParams.alpha;
for jj=1:nSamples
    db=1i*4*pi^2*(f1(jj)-f).*(f2-f)*fiberParams.GVD;
    p1=abs((1-exp(da*sys.Ls).*exp(db*sys.Ls))./(-da-db)).^2;
    if sys.Ns==1
        chi=ones(size(db));
    else
        den=sin(db/2*sys.Ls);
        chi=abs(sin(db/2*sys.Ns*sys.Ls)./den).^2;
        chi(abs(den)<eps)=sys.Ns^2;
    end
    shift=index-jj;
    if shift>=0
        g3=[zeros(1,shift),gtx(1:end-shift)];
    else
        s=abs(shift); g3=[gtx(1+s:end),zeros(1,s)];
    end
    I2(jj)=trapz(f2,p1.*chi.*gtx(jj).*gtx.*g3);
end
I1=trapz(f1,I2);
GNLI=(16/27)*fiberParams.gamma^2*I1*1e-12;
end
