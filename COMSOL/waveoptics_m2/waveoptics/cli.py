"""Command-line user tool: solve, plot, and export numeric data."""
import argparse
import csv
import json
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from .slab import solve_slab


def main(argv=None):
    parser = argparse.ArgumentParser(description='M1: 1D symmetric slab TE/TM P1 FEM mode solver (lengths in um).')
    for name,default in [('n-core',1.5),('n-clad',1.45),('width-um',3.0),
                         ('wavelength-um',1.55),('padding-um',20.0),('h-um',0.0375)]:
        parser.add_argument('--'+name,type=float,default=default)
    parser.add_argument('--polarization',choices=['TE','TM','both'],default='both')
    parser.add_argument('--max-modes',type=int,default=6)
    parser.add_argument('--mesh-backend',choices=['gmsh','native'],default='gmsh')
    parser.add_argument('--out',type=Path,default=Path('slab_output'))
    args = parser.parse_args(argv)
    params = vars(args).copy()
    out = params.pop('out')
    polarization = params.pop('polarization')
    out.mkdir(parents=True,exist_ok=True)
    summaries = []
    pols = ['TE','TM'] if polarization == 'both' else [polarization]
    fig,axes = plt.subplots(len(pols),1,figsize=(9,3.5*len(pols)),squeeze=False)
    for ax,pol in zip(axes[:,0],pols):
        sol = solve_slab(**params,polarization=pol)
        header = ['x_um']+[f'{pol}{m.order}_scalar_L2_normalized' for m in sol.modes]
        data = np.column_stack([sol.x_um]+[m.field for m in sol.modes])
        np.savetxt(out/f'fields_{pol}.csv',data,delimiter=',',header=','.join(header),comments='')
        a = params['width_um']/2
        ax.axvspan(-a,a,color='#dce8f0',label='core')
        for mode in sol.modes:
            row = {'mode':f'{pol}{mode.order}','neff':mode.neff,'beta_per_um':mode.beta_per_um,
                   'relative_residual':mode.relative_residual,'boundary_decay_estimate':mode.boundary_decay_estimate}
            summaries.append(row)
            print(f"{row['mode']}: n_eff={mode.neff:.12f}, beta={mode.beta_per_um:.12f} /um, residual={mode.relative_residual:.3e}")
            ax.plot(sol.x_um,mode.field,label=f"{pol}{mode.order}: n_eff={mode.neff:.9f}")
        ax.set(xlabel='x (um)',ylabel=f'{pol}: '+('Ey' if pol=='TE' else 'Hy')+' (scalar L2 normalized)',
               title=f'{pol} | {sol.elements} elements | h_max={sol.h_max_um:.6g} um')
        ax.legend(fontsize=8)
        ax.grid(alpha=0.2)
    fig.tight_layout()
    fig.savefig(out/'modes.png',dpi=180)
    plt.close(fig)
    if summaries:
        with (out/'modes.csv').open('w',newline='') as f:
            writer = csv.DictWriter(f,fieldnames=list(summaries[0]))
            writer.writeheader()
            writer.writerows(summaries)
    (out/'run.json').write_text(json.dumps({'parameters':params,'polarization':polarization,'modes':summaries},indent=2)+'\n')


if __name__ == '__main__':
    main()
