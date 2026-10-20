"""Auditable raster sampling of Zhao et al. Fig.2a; approximate data only.
Preserve source hash, axis calibration and crop. Never treat as raw COMSOL export.
"""
from pathlib import Path
import urllib.request,subprocess,hashlib,json
import numpy as np
from PIL import Image
ROOT=Path(__file__).resolve().parent;OUT=ROOT/'results'
URL='https://content.scirp.org/pdf/opj_2021082710523279.pdf'
def runs(mask):
    d=np.diff(np.r_[False,mask,False].astype(int));return list(zip(np.flatnonzero(d==1),np.flatnonzero(d==-1)))
def longest(mask):
    r=runs(mask);return max(r,key=lambda z:z[1]-z[0]) if r else (0,0)
def main():
    out=dict(source_url=URL,source_locator='Figure 2(a), printed p.405, PDF page 4',
        method='green FEM curve sampled from raster; not raw numerical COMSOL data',
        axis_assumptions=dict(gap_min_um=0,gap_max_um=10,kappa_min_per_m=0,kappa_max_per_m=1600),
        status='unavailable')
    try:
        pdf=OUT/'mcf_source_temp.pdf';urllib.request.urlretrieve(URL,pdf)
        out['source_pdf_sha256']=hashlib.sha256(pdf.read_bytes()).hexdigest()
        subprocess.run(['pdftoppm','-f','4','-l','4','-singlefile','-r','180','-png',str(pdf),str(OUT/'mcf_figure_temp')],check=True,capture_output=True)
        image=Image.open(OUT/'mcf_figure_temp.png').convert('RGB');a=np.asarray(image)/255.
        H,W=a.shape[:2];green=(a[:,:,1]>a[:,:,0]*1.35)&(a[:,:,1]>a[:,:,2]*1.35)&(a[:,:,1]>.15)
        region=np.zeros((H,W),bool);region[int(.06*H):int(.46*H),int(.08*W):int(.66*W)]=True
        yy,xx=np.where(green&region)
        if len(xx)<20:raise RuntimeError('Insufficient green curve pixels')
        dark=np.max(a,axis=2)<.30
        ylo=max(int(.06*H),int(yy.min())-45);yhi=min(int(.46*H),int(yy.max())+60)
        candidates=[]
        for y in range(ylo,yhi):
            left,right=longest(dark[y,int(.08*W):int(.66*W)])
            left+=int(.08*W);right+=int(.08*W)
            if right-left>.30*W:candidates.append((y,left,right))
        if not candidates:raise RuntimeError('Plot horizontal borders not found')
        bottom,left,right=max(candidates)
        xaxis=left
        # The left and right borders have a continuous vertical dark run.
        possible=[]
        for x in range(left-2,left+4):
            v0,v1=longest(dark[ylo:bottom+1,x])
            if v1-v0>.10*H:possible.append((v1-v0,ylo+v0,ylo+v1,x))
        if not possible:raise RuntimeError('Plot vertical border not found')
        _,top,_,xaxis=max(possible)
        if not (top<bottom and right-left>.30*W):raise RuntimeError('Axis calibration invalid')
        out['calibration_pixels']=dict(left=int(xaxis),right=int(right-1),top=int(top),bottom=int(bottom),
            raster_width=W,raster_height=H,assumption='horizontal box spans dg 0..10 um; vertical box spans kappa 0..1.6e-3 per um')
        samples=[]
        for gap in [1.,2.,4.,6.]:
            x=xaxis+(right-1-xaxis)*gap/10
            mask=(green[:,max(0,int(round(x))-2):int(round(x))+3])
            ys=np.where(mask[top+2:bottom-2].any(axis=1))[0]+top+2
            if len(ys)==0:continue
            # Legend is at high x; sample y uses median of colored curve pixels.
            y=float(np.median(ys));k=1600*(bottom-y)/(bottom-top)
            unc=1600*(max(2.,float(ys.max()-ys.min())/2)+2)/(bottom-top)
            samples.append(dict(gap_um=gap,kappa_estimate_per_m=float(k),raster_uncertainty_per_m=float(unc),
                x_pixel=float(x),y_pixel=y))
        if not any(r['gap_um']==2 for r in samples):raise RuntimeError('Base dg=2 point not found')
        out.update(status='estimated',samples=samples,
            limitation='Raster uncertainty is approximate; it excludes axis/convention/model ambiguity and undocumented COMSOL mesh or polarization.')
        image.crop((max(0,xaxis-80),max(0,top-45),min(W,right+30),min(H,bottom+65))).save(OUT/'mcf_published_figure2a.png')
        print('FIGURE_TEXT '+subprocess.check_output(['pdftotext','-f','4','-l','4','-layout',str(pdf),'-'],text=True)[:5000],flush=True)
    except Exception as e:out['error']=str(e)
    finally:
        for name in ['mcf_source_temp.pdf','mcf_figure_temp.png']:
            (OUT/name).unlink(missing_ok=True)
    (OUT/'mcf_graph_estimates.json').write_text(json.dumps(out,indent=2)+'\n')
    print('MCF_FIGURE_ESTIMATES '+json.dumps(out),flush=True)
if __name__=='__main__':main()
