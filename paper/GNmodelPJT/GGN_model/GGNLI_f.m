function GGf = GGNLI_f(fiberParams, signal, sys, f, fibProf, sigPSD, zeta_samples)
%GGNLI_F Matrix-style generalized GN model with arbitrary power profiles.
if nargin<7 || isempty(zeta_samples), zeta_samples=size(fibProf,2); end
nSamples=numel(sigPSD); grid=linspace(-sys.Bw/2,sys.Bw/2,nSamples);
z=linspace(0,sys.Ls,zeta_samples);
if size(fibProf,2)~=zeta_samples, error('fibProf columns must equal zeta_samples'); end
prof=zeros(nSamples,zeta_samples);
for iz=1:zeta_samples
    prof(:,iz)=interp1(sys.fCh,fibProf(:,iz),grid,'linear','extrap');
end
prof=max(prof,realmin); [~,idx]=min(abs(grid-f)); I2=zeros(1,nSamples);
for j=1:nSamples
    k=(1:nSamples)+(idx-j); valid=k>=1 & k<=nSamples;
    if ~any(valid) || sigPSD(j)==0, continue; end
    kv=k(valid); f2=grid(valid);
    db=1i*4*pi^2*(grid(j)-f).*(f2-f)*fiberParams.GVD;
    ratio=sqrt(prof(j,:).*prof(valid,:)./prof(kv,:)).*sqrt(prof(idx,:));
    phase=exp(db(:).*z); F=abs(trapz(z,phase.*ratio,2)).^2;
    I2(j)=trapz(f2,F.'.*sigPSD(j).*sigPSD(valid).*sigPSD(kv));
end
GGf=(16/27)*fiberParams.gamma^2*trapz(grid,I2)*1e-12;
end
