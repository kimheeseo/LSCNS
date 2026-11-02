"""Fast architectural smoke test: no official OFC CSV or misleading performance scores."""
import torch
from model import PIEDFAFNO

def main():
    torch.manual_seed(2026)
    model=PIEDFAFNO(global_dim=10,pi_hidden=96,width=32,modes=12,depth=2)
    model.freeze_physics()
    n=8
    g=torch.randn(n,10)
    mask=(torch.rand(n,95)>.3).float()
    chan=torch.randn(n,4,95)
    chan[:,0,:]=mask
    share=mask / mask.sum(dim=1,keepdim=True).clamp_min(1)
    target=torch.full((n,),20.)
    tilt=torch.zeros(n)
    features=(g,chan,mask,share,target,tilt)
    baseline,_=model.physics(*features)
    pred,components=model(*features)
    assert pred.shape==(n,95)
    assert torch.allclose(pred,baseline,atol=1e-6), 'Zero-init residual mismatch'
    assert (components['local_correction'][mask<.5]==0).all(), 'Inactive residual'
    goal=baseline.detach()+.2*torch.sin(torch.linspace(0,3.14,95))[None,:]
    opt=torch.optim.Adam(model.residual_net.parameters(),lr=.001)
    opt.zero_grad()
    loss=(((pred-goal)*mask)**2).mean()
    loss.backward()
    assert any(p.grad is not None and p.grad.abs().sum()>0
               for p in model.residual_net.parameters()), 'FNO gradients missing'
    opt.step()
    pred2,_=model(*features)
    assert torch.isfinite(pred2).all()
    print('PASS: shape, zero residual, inactive mask, FFT gradients, optimizer step.')
    print('This is a synthetic functional test, NOT a Kaggle/OFC performance result.')

if __name__=='__main__':main()
