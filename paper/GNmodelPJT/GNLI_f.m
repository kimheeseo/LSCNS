function GNLI = GNLI_f(fiberParams, signal, sys, f, gtx)
nSamples=floor(sys.Bw*23000/8);
f2=linspace(-sys.Bw/2,sys.Bw/2,nSamples);
f1=linspace(-sys.Bw/2,sys.Bw/2,nSamples);
da=-2*fiberParams.alpha;
gtxf2=gtx;
gtxf1=gtxf2;
I2=zeros(1,nSamples);
n=nSamples;
[~,index]=min(abs(f-f2));
for jj=1:n
    db=1i*4*pi^2*(f1(jj)-f)*(f2-f)*fiberParams.GVD;
    p1=abs((1-exp(da*sys.Ls).*exp(db*sys.Ls))./(-da-db)).^2;
    if sys.Ns==1
        chi=ones(1,nSamples);
    else
        chi=(sin(db./2*sys.Ns*sys.Ls)./sin(db./2*sys.Ls)).^2;
    end
    vec2=find(db==0);
    chi(vec2)=sys.Ns^2;
    if index-jj>=0
        ze=zeros(1,index-jj);
        gtxf3=horzcat(ze,gtxf2(1:(n-length(ze))));
        I2(jj)=trapz(f2,p1.*chi.*gtxf1(jj).*gtxf2.*gtxf3);
    else
        ze=zeros(1,abs(index-jj));
        gtxf3=horzcat(gtxf2(1+length(ze):end),ze);
        I2(jj)=trapz(f2,p1.*chi.*gtxf1(jj).*gtxf2.*gtxf3);
    end
    if mod(jj,100)==0
        fprintf('%g%%\n',round(jj*100/nSamples));
    end
end
save('FUNCTION.mat','I2');
I1=trapz(f1,I2);
GNLI=16/27*fiberParams.gamma^2*I1*1e-12;
end
