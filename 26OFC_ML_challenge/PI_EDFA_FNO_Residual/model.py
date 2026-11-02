"""PI-EDFA + 1D Fourier Neural Operator residual for 95 optical channels."""
from __future__ import annotations
import sys
from pathlib import Path
import torch
from torch import nn
import torch.nn.functional as F

BASE = Path(__file__).resolve().parent.parent
PI_DIR = BASE / "COSMOS_Transfer_Learning"
if str(PI_DIR) not in sys.path:
    sys.path.insert(0, str(PI_DIR))
import physics_edfa as pi
N_CH = pi.N_CH

class SpectralConv1d(nn.Module):
    """Truncated Fourier multiplier with learnable complex channel mixing."""
    def __init__(self, in_channels, out_channels, modes=16):
        super().__init__()
        if not 1 <= modes <= N_CH // 2 + 1:
            raise ValueError("modes out of range")
        self.out_channels, self.modes = out_channels, modes
        self.weight = nn.Parameter(torch.randn(in_channels, out_channels, modes, 2) *
                                   (1/(in_channels*out_channels)**.5))

    def forward(self, x):
        freq = torch.fft.rfft(x.float(), dim=-1, norm="ortho")
        w = torch.view_as_complex(self.weight.contiguous())
        k = min(self.modes, freq.size(-1))
        out = torch.zeros(x.size(0), self.out_channels, freq.size(-1),
                          device=x.device, dtype=freq.dtype)
        out[:, :, :k] = torch.einsum("bim,iom->bom", freq[:, :, :k], w[:, :, :k])
        return torch.fft.irfft(out, n=x.size(-1), dim=-1, norm="ortho").to(x.dtype)

class FNOBlock(nn.Module):
    def __init__(self, width, modes, global_dim):
        super().__init__()
        self.spectral = SpectralConv1d(width, width, modes)
        self.local = nn.Conv1d(width, width, 1)
        self.condition = nn.Linear(global_dim, 2*width)
        self.norm = nn.GroupNorm(1, width)
        nn.init.zeros_(self.condition.weight)
        nn.init.zeros_(self.condition.bias)

    def forward(self, x, g):
        v = self.norm(self.spectral(x)+self.local(x))
        scale, shift = self.condition(g).chunk(2, dim=-1)
        v = v*(1+.1*torch.tanh(scale)[:,:,None])+shift[:,:,None]
        return x+F.gelu(v)

class FNOResidual(nn.Module):
    """No labels, category IDs, or target test measurements as inputs."""
    def __init__(self, global_dim=10, width=48, modes=16, depth=3):
        super().__init__()
        self.project = nn.Conv1d(6, width, 1)
        self.global_project = nn.Linear(global_dim, width)
        self.blocks = nn.ModuleList([FNOBlock(width,modes,global_dim)
                                     for _ in range(depth)])
        self.head = nn.Sequential(nn.Conv1d(width,width,1),nn.GELU(),nn.Conv1d(width,1,1))
        nn.init.zeros_(self.head[-1].weight)
        nn.init.zeros_(self.head[-1].bias)

    def forward(self, global_x, channel_x, mask, share, baseline):
        raw = torch.cat([channel_x, (baseline/25)[:,None,:],
                         (share*N_CH)[:,None,:]], dim=1)
        x = self.project(raw)+self.global_project(global_x)[:,:,None]
        for block in self.blocks: x = block(x, global_x)
        return self.head(x).squeeze(1)*mask

class PIEDFAFNO(nn.Module):
    """Predicted Gain = frozen PI-EDFA output + Fourier spectral correction."""
    def __init__(self,global_dim=10,pi_hidden=96,width=48,modes=16,depth=3):
        super().__init__()
        self.physics = pi.PhysicalSpectralNet(global_dim=global_dim,hidden=pi_hidden)
        self.residual_net = FNOResidual(global_dim,width,modes,depth)

    def freeze_physics(self):
        for p in self.physics.parameters():p.requires_grad_(False)
        self.physics.eval()

    def forward(self,global_x,channel_x,mask,share,target_gain,target_tilt):
        if not any(p.requires_grad for p in self.physics.parameters()):
            with torch.no_grad():
                baseline,_ = self.physics(global_x,channel_x,mask,share,
                                          target_gain,target_tilt)
        else:
            baseline,_ = self.physics(global_x,channel_x,mask,share,
                                      target_gain,target_tilt)
        correction=self.residual_net(global_x,channel_x,mask,share,baseline.detach())
        return baseline+correction, {"base":baseline,"local_correction":correction}
