"""Batch interface for M2; fields.npz retains unsmoothed edge/scalar DOFs."""
import argparse
import json
from pathlib import Path
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from .vector import fiber_mesh,rectangle_mesh,solve_vector


def main():
    parser = argparse.ArgumentParser(description='Full-vector Nedelec/P1 waveguide modes (PEC exterior)')
    parser.add_argument('geometry',choices=['fiber','rectangle'])
    parser.add_argument('--wavelength',type=float,default=1.)
    parser.add_argument('--n-core',type=complex,default=1.5)
    parser.add_argument('--n-clad',type=complex,default=1.)
    parser.add_argument('--core-radius',type=float,default=.3)
    parser.add_argument('--outer-radius',type=float,default=3.)
    parser.add_argument('--width',type=float,default=1.2)
    parser.add_argument('--height',type=float,default=.9)
    parser.add_argument('--h',type=float,default=.03)
    parser.add_argument('--target',type=float,default=1.25)
    parser.add_argument('--window',type=float,nargs=2,default=(1.0001,1.4999))
    parser.add_argument('--candidates',type=int,default=8)
    parser.add_argument('--output',type=Path,default=Path('vector_results'))
    args = parser.parse_args()
    if args.geometry == 'fiber':
        mesh = fiber_mesh(args.core_radius,args.outer_radius,args.h)
        n = np.where(mesh.region==1,args.n_core,args.n_clad)
    else:
        mesh = rectangle_mesh(args.width,args.height,args.h)
        n = args.n_core
    sol = solve_vector(mesh,n,wavelength_um=args.wavelength,target_neff=args.target,
                      neff_window=args.window,candidates=args.candidates)
    args.output.mkdir(parents=True,exist_ok=True)
    sol.save_npz(args.output/'fields.npz')
    rows = [dict(index=i,neff_real=float(m.neff.real),neff_imag=float(m.neff.imag),
                 effective_area_um2=m.effective_area_um2,loss_db_per_m=m.loss_db_per_m,
                 pencil_residual=m.pencil_residual,gauss_residual=m.gauss_residual)
            for i,m in enumerate(sol.modes)]
    data = dict(modes=rows,diagnostics=sol.diagnostics,geometry=mesh.metadata)
    (args.output/'summary.json').write_text(json.dumps(data,indent=2,default=lambda x:x.item()))
    for i,m in enumerate(sol.modes):
        et = sol.edge_basis.interpolate(m.edge_coefficients)
        ez = -1j*m.beta_per_um*sol.scalar_basis.interpolate(m.scalar_coefficients)
        intensity = np.sum(abs(et)**2,axis=0)+abs(ez)**2
        cell_intensity = np.sum(intensity*sol.edge_basis.dx,axis=1)/np.sum(sol.edge_basis.dx,axis=1)
        fig,ax = plt.subplots(figsize=(6,5))
        im = ax.tripcolor(mesh.mesh.p[0],mesh.mesh.p[1],mesh.mesh.t.T,
                          facecolors=cell_intensity,shading='flat',cmap='magma')
        ax.set(aspect='equal',xlabel='x (um)',ylabel='y (um)',title=f'Mode {i}  n_eff={m.neff.real:.9f}{m.neff.imag:+.3e}i')
        fig.colorbar(im,ax=ax,label='Cell-averaged |E|² (normalized)')
        fig.tight_layout();fig.savefig(args.output/f'mode_{i}.png',dpi=180);plt.close(fig)
    print(json.dumps(data,indent=2,default=lambda x:x.item()))


if __name__ == '__main__':
    main()
