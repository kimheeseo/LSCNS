function fiberParams = fiberParams(fiberType)
%FIBERPARAMS Return fiber parameters used by the GN/GGN examples.
% alpha: field attenuation [1/km], GVD: beta2 [ps^2/km], gamma [1/W/km].
if nargin < 1, fiberType = 'SSMF'; end
switch upper(string(fiberType))
    case "SSMF"
        alphaDb = 0.20; D = 16.7; gamma = 1.30;
    case "NZDSF"
        alphaDb = 0.22; D = 4.5; gamma = 1.50;
    otherwise
        error('Unknown fiber type: %s', fiberType);
end
c_nm_ps = 299792.458; lambda_nm = 1550;
beta2 = -(D*lambda_nm^2)/(2*pi*c_nm_ps);
fiberParams = struct('alpha', alphaDb*log(10)/20, ...
    'GVD', beta2, 'gamma', gamma, 'D', D, 'alphaDb', alphaDb);
end
