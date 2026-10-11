"""Independent closed-form M3 oracles. No waveoptics imports.

Time convention exp(-i*w*t), outgoing Hankel(1); lengths um.
Layer transfer matrices follow continuity of u and p*du/dx.
Cylinder coefficients follow continuity of u and p*du/dr at r=a.
"""
import numpy as np
from scipy.special import jv, jvp, hankel1, h1vp


def guide_beta(n, wavelength, width, order):
    return np.sqrt((2*np.pi*n/wavelength)**2-(order*np.pi/width)**2+0j)


def layer_s(n_list, lengths, wavelength, width, order, polarization):
    k0=2*np.pi/wavelength
    betas=np.array([guide_beta(n,wavelength,width,order) for n in n_list])
    ps=np.ones(len(n_list)) if polarization=='TE' else 1/np.array(n_list,dtype=complex)**2
    ys=ps*betas
    transfer=np.eye(2,dtype=complex)
    # state [u, p*du/dx] at right = transfer @ state at left
    for beta,p,length in zip(betas,ps,lengths):
        co,si=np.cos(beta*length),np.sin(beta*length)
        transfer=np.array([[co,si/(p*beta)],[-p*beta*si,co]])@transfer
    vin=np.array([1,1j*ys[0]])
    vref=np.array([1,-1j*ys[0]])
    vout=np.array([1,1j*ys[-1]])
    # transfer @ (vin+r*vref) = t*vout
    r,t=np.linalg.solve(np.column_stack([transfer@vref,-vout]),-transfer@vin)
    return complex(r),complex(t*np.sqrt(ys[-1].real/ys[0].real))


def cylinder_coefficients(radius,n_inside,n_background,wavelength,polarization,orders=24):
    k0=2*np.pi/wavelength; ki=k0*n_inside; kb=k0*n_background
    pi,pb=(1.,1.) if polarization=='TE' else (1/n_inside**2,1/n_background**2)
    m=np.arange(-orders,orders+1)
    Ji=jv(m,ki*radius); Jb=jv(m,kb*radius); Hb=hankel1(m,kb*radius)
    numerator=pi*ki*jvp(m,ki*radius)*Jb-pb*kb*jvp(m,kb*radius)*Ji
    denominator=pb*kb*h1vp(m,kb*radius)*Ji-pi*ki*jvp(m,ki*radius)*Hb
    return m,numerator/denominator


def cylinder_field(xy,radius,n_inside,n_background,wavelength,polarization,orders=24):
    xy=np.asarray(xy); r=np.hypot(xy[0],xy[1]); theta=np.arctan2(xy[1],xy[0])
    if np.any(r<=radius): raise ValueError('Exterior scattered-field oracle only.')
    m,c=cylinder_coefficients(radius,n_inside,n_background,wavelength,polarization,orders)
    return np.sum((1j**m*c)[:,None]*hankel1(m[:,None],2*np.pi*n_background/wavelength*r)*np.exp(1j*m[:,None]*theta),axis=0)


def cylinder_width(radius,n_inside,n_background,wavelength,polarization,orders=24):
    _,c=cylinder_coefficients(radius,n_inside,n_background,wavelength,polarization,orders)
    return float(4/(2*np.pi*n_background/wavelength)*np.sum(np.abs(c)**2))
