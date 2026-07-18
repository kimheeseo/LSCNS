function rho = fiber_var(power, Cr, fMax, alpha, deltaf, Ls, n, nCh, J, zeta_samples)
%FIBER_VAR Raman-tilted normalized power evolution rho(z,f)^2.
m = Cr/fMax; G = m*deltaf; z = linspace(0,Ls,zeta_samples);
Leff=(1-exp(-2*alpha(n).*z))./(2*alpha(n));
A=zeros(nCh,numel(z));
for ii=1:nCh
    A(ii,:)=power(ii).*exp(G*J*(ii-1).*Leff);
end
A=sum(A,1);
rho=J.*exp(-2*alpha(n).*z).*exp(G*J*(n-1).*Leff)./A;
end
