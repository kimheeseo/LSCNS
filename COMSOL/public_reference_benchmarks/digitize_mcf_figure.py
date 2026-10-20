"""Approximate raster sampling of Zhao et al. Fig.2(a), PDF p.405.
Axis coordinates were checked by viewing the actual 180-dpi rendering.
Red = FEM, blue = CMT. This is not raw COMSOL numerical output.
"""
from pathlib import Path
import urllib.request,subprocess,hashlib,json
import numpy as np
from PIL import Image
ROOT=Path(__file__).resolve().parent;OUT=ROOT/'results'
URL='https://content.scirp.org/pdf/opj_2021082710523279.pdf'
EXPECTED_PDF_SHA256='b82f0faf6b0036fe2be9f49fb967c2ca320b664e8e2b6169e45feb18bdec82d9'

def main():
    out=dict(source_url=URL,source_locator='Figure 2(a), printed p.405, PDF page 4',
        method='red FEM curve, manually checked axes; approximate raster sampling, not raw COMSOL',
        status='unavailable',attribution='Zhao, Ren, Yin and Wang (2021), DOI 10.4236/opj.2021.118029, CC BY 4.0; crop and sampling by this project')
    try:
        pdf=OUT/'mcf_source_temp.pdf';errors=[]
        for candidate in [URL,'https://www.scirp.org/pdf/opj_2021082710523279.pdf','https://www.scirp.net/pdf/opj_2021082710523279.pdf']:
            try:
                req=urllib.request.Request(candidate,headers={'User-Agent':'Mozilla/5.0 (compatible; public-research-reference-comparison)','Accept':'application/pdf'})
                with urllib.request.urlopen(req,timeout=30) as response:body=response.read()
                if not body.startswith(b'%PDF'):raise RuntimeError('Not a PDF response')
                pdf.write_bytes(body);out['retrieved_url']=candidate;break
            except Exception as err:errors.append(str(err))
        if not pdf.exists():raise RuntimeError('; '.join(errors))
        sha=hashlib.sha256(pdf.read_bytes()).hexdigest();out['source_pdf_sha256']=sha
        if sha!=EXPECTED_PDF_SHA256:raise RuntimeError('Source PDF changed; manual recalibration required')
        subprocess.run(['pdftoppm','-f','4','-l','4','-singlefile','-r','180','-png',str(pdf),str(OUT/'mcf_figure_temp')],check=True,capture_output=True)
        image=Image.open(OUT/'mcf_figure_temp.png').convert('RGB');a=np.asarray(image)/255.;H,W=a.shape[:2]
        if W!=1489 or not 2018<=H<=2023:raise RuntimeError('Unexpected raster dimensions; calibration not applicable')
        # Coordinates observed on exact source above: x=1..30 um, y=0..2000 /m.
        left,right,top,bottom=691.,1223.,186.,425.
        out['axis_calibration']=dict(left_pixel=left,right_pixel=right,top_pixel=top,bottom_pixel=bottom,
            gap_min_um=1.,gap_max_um=30.,kappa_min_per_m=0.,kappa_max_per_m=2000.,dpi=180,
            verification='axes/legend manually inspected; exact PDF hash and fixed rendering resolution')
        red=(a[:,:,0]>.55)&(a[:,:,0]>a[:,:,1]*1.5)&(a[:,:,0]>a[:,:,2]*1.5)
        samples=[]
        for gap in [1.,2.,4.,6.]:
            x=left+(right-left)*(gap-1.)/29.
            lo,hi=int(round(x))-2,int(round(x))+3
            ys=np.where(red[int(top)+3:int(bottom)-3,lo:hi].any(axis=1))[0]+int(top)+3
            if not len(ys):raise RuntimeError('No red FEM pixels near gap '+str(gap))
            y=float(np.median(ys));k=2000*(bottom-y)/(bottom-top)
            # Colored thickness + +/-2 pixels of x/axis rounding. Not statistical CI.
            uncertainty=2000*(max(2.,float(ys.max()-ys.min())/2)+2)/(bottom-top)
            samples.append(dict(gap_um=gap,kappa_estimate_per_m=float(k),raster_uncertainty_per_m=float(uncertainty),
                x_pixel=float(x),y_pixel=y))
        out.update(status='estimated',samples=samples,
            limitation='Approximate raster uncertainty, not a confidence interval. Does not include Delta convention, polarization, model or COMSOL discretization uncertainty.')
        image.crop((int(left)-80,int(top)-40,int(right)+25,int(bottom)+70)).save(OUT/'mcf_published_figure2a.png')
    except Exception as err:out['error']=str(err)
    finally:
        for name in ['mcf_source_temp.pdf','mcf_figure_temp.png']:(OUT/name).unlink(missing_ok=True)
    (OUT/'mcf_graph_estimates.json').write_text(json.dumps(out,indent=2)+'\n')
    print('MCF_FIGURE_ESTIMATES '+json.dumps(out),flush=True)
if __name__=='__main__':main()
