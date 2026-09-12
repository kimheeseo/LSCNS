function rho = fiber(power, Cr, fMax, alpha, deltaf, Ls, n, nCh, J)
%FIBER Raman-tilted normalized power at the end of one span.
m=Cr/fMax; G=m*deltaf;
Leff=(1-exp(-2*alpha*Ls))/(2*alpha);
A=0;
for ii=1:nCh
    A=A+power(ii)*exp(G*J*(ii-1)*Leff);
end
rho=J*exp(-2*alpha*Ls)*exp(G*J*(n-1)*Leff)/A;
end
