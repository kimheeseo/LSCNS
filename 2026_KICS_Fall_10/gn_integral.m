function out = gn_integral_math_verification()
% Mathematical verification helpers for Poggiolini GNRF Eq. (G.4).
% Units: f [THz], beta2 [ps^2/km], L [km], gamma [1/(W km)].
% alpha_dB is conventional POWER loss; field alpha = alpha_dB*ln(10)/20.

alpha_dB = 0.2; D = 17; lambda_nm = 1550; L = 80; N = 10;
alpha = alpha_dB*log(10)/20;
c_nm_ps = 299792.458;
beta2 = -D*lambda_nm^2/(2*pi*c_nm_ps);
Leff = (1-exp(-2*alpha*L))/(2*alpha);

q = 4*pi^2*beta2*(0.016)*( -0.016 );
Hs_complex = abs((1-exp((-2*alpha+1i*q)*L))/(2*alpha-1i*q))^2;
e = exp(-2*alpha*L);
Hs_stable = (1+e^2-2*e*cos(q*L))/((2*alpha)^2+q^2);

x = q*L/2;
PA = (sin(N*x)/sin(x))^2;
PA0 = N^2;
Hs0 = Leff^2;

P1 = 1e-3; P2 = 2e-3; B = 0.032; gamma = 1.3;
G1 = P1/B; G2 = P2/B;
I1 = (16/27)*gamma^2*G1^3*Hs_stable*PA;
I2 = (16/27)*gamma^2*G2^3*Hs_stable*PA;
ratio = I2/I1;

out.alpha_field_1_per_km = alpha;
out.beta2_ps2_per_km = beta2;
out.Leff_km = Leff;
out.span_equation_error_pct = abs(Hs_complex-Hs_stable)/Hs_complex*100;
out.span_limit_error_pct = abs(Hs0-Leff^2)/(Leff^2)*100;
out.phase_limit_error_pct = abs(PA0-N^2)/(N^2)*100;
out.power_cubic_ratio = ratio;
out.power_cubic_error_pct = abs(ratio-8)/8*100;

disp(out)
end
