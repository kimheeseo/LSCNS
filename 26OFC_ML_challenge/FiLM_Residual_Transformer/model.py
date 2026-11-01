"""95-channel EDFA gain predictor: analytical residual + per-block FiLM + SHB-aware curvature loss.

Uses only inference-available conditioning variables. Scalers must be fitted on the training split.
"""
from __future__ import annotations

import torch
from torch import nn
import torch.nn.functional as F


class AnalyticalGainBaseline(nn.Module):
    """Approximate gain = gain setpoint + linear tilt - mild input-shape correction.

    kappa is an empirical prior, NOT an erbium rate-equation solution.
    """
    def __init__(self, n_channels: int = 95, kappa: float = 0.02):
        super().__init__()
        self.kappa = kappa
        self.register_buffer("position", torch.linspace(-1.0, 1.0, n_channels))

    def forward(self, pin_dbm, active_mask, target_gain, target_tilt):
        on = active_mask.float()
        pin = torch.nan_to_num(pin_dbm, nan=-80.0, posinf=20.0, neginf=-80.0)
        pin = pin.clamp(-80.0, 20.0)
        avg_pin = (pin * on).sum(dim=1, keepdim=True) / on.sum(dim=1, keepdim=True).clamp_min(1.0)
        pin_deviation = ((pin - avg_pin) * on).clamp(-30.0, 30.0)
        # Target tilt = left-to-right end-to-end gain difference (dB).
        return (target_gain[:, None]
                + 0.5 * target_tilt[:, None] * self.position[None, :]
                - self.kappa * pin_deviation)


class FiLM(nn.Module):
    """Condition-wise scale and shift: x' = (1 + gamma(c)) * x + beta(c)."""
    def __init__(self, condition_dim: int, hidden_dim: int):
        super().__init__()
        self.affine = nn.Linear(condition_dim, 2 * hidden_dim)
        # Identity at initialization, stable for residual learning.
        nn.init.zeros_(self.affine.weight)
        nn.init.zeros_(self.affine.bias)

    def forward(self, token_features, condition):
        gamma, beta = self.affine(condition).chunk(2, dim=-1)
        return token_features * (1.0 + gamma[:, None, :]) + beta[:, None, :]


class FiLMTransformerBlock(nn.Module):
    """FiLM is applied *inside* attention and feed-forward branches."""
    def __init__(self, hidden_dim: int = 64, n_heads: int = 4, condition_dim: int = 3,
                 dropout: float = 0.08):
        super().__init__()
        self.attn_norm = nn.LayerNorm(hidden_dim)
        self.attn_film = FiLM(condition_dim, hidden_dim)
        self.attn = nn.MultiheadAttention(hidden_dim, n_heads, dropout=dropout, batch_first=True)
        self.ffn_norm = nn.LayerNorm(hidden_dim)
        self.ffn_film = FiLM(condition_dim, hidden_dim)
        self.ffn = nn.Sequential(nn.Linear(hidden_dim, 2 * hidden_dim), nn.GELU(),
                                 nn.Dropout(dropout), nn.Linear(2 * hidden_dim, hidden_dim))
        self.dropout = nn.Dropout(dropout)

    def forward(self, tokens, condition):
        q = self.attn_film(self.attn_norm(tokens), condition)
        update, _ = self.attn(q, q, q, need_weights=False)
        tokens = tokens + self.dropout(update)
        h = self.ffn_film(self.ffn_norm(tokens), condition)
        return tokens + self.dropout(self.ffn(h))


class ResidualFiLMEDFA(nn.Module):
    def __init__(self, cond_mean, cond_std, n_channels: int = 95,
                 hidden_dim: int = 64, n_heads: int = 4, n_layers: int = 2):
        super().__init__()
        assert n_channels == 95, "This OFC implementation expects exactly 95 channels."
        self.register_buffer("cond_mean", torch.as_tensor(cond_mean, dtype=torch.float32))
        self.register_buffer("cond_std", torch.as_tensor(cond_std, dtype=torch.float32).clamp_min(1e-3))
        self.baseline = AnalyticalGainBaseline(n_channels=n_channels)
        self.register_buffer("position", torch.linspace(-1.0, 1.0, n_channels))
        self.embed = nn.Linear(4, hidden_dim)
        self.blocks = nn.ModuleList([
            FiLMTransformerBlock(hidden_dim, n_heads, condition_dim=3)
            for _ in range(n_layers)
        ])
        self.head_norm = nn.LayerNorm(hidden_dim)
        self.residual_head = nn.Linear(hidden_dim, 1)
        nn.init.zeros_(self.residual_head.weight)
        nn.init.zeros_(self.residual_head.bias)

    def config(self) -> dict:
        """Reconstruction metadata stored beside trained state_dict."""
        return {
            "cond_mean": self.cond_mean.cpu().tolist(),
            "cond_std": self.cond_std.cpu().tolist(),
            "n_channels": len(self.position),
            "hidden_dim": self.embed.out_features,
            "n_heads": self.blocks[0].attn.num_heads,
            "n_layers": len(self.blocks),
        }

    def forward(self, pin_dbm, active_mask, aging, target_gain, target_tilt):
        """pin [B,95] dBm; mask [B,95] 0/1; scalars [B]."""
        on = active_mask.float()
        pin = torch.nan_to_num(pin_dbm, nan=-80.0, posinf=20.0, neginf=-80.0).clamp(-80.0, 20.0)
        pin_mw = (10.0 ** (pin / 10.0)) * on  # dBm -> mW, then mask
        share = pin_mw / pin_mw.sum(dim=1, keepdim=True).clamp_min(1e-12)
        avg_pin = (pin * on).sum(dim=1, keepdim=True) / on.sum(dim=1, keepdim=True).clamp_min(1.0)
        local_shape = ((pin - avg_pin) * on / 10.0).clamp(-3.0, 3.0)
        position = self.position[None, :].expand_as(on)
        tokens = torch.stack([local_shape, torch.log1p(95.0 * share), on, position], dim=-1)
        tokens = self.embed(tokens)

        raw_cond = torch.stack([aging, target_gain, target_tilt], dim=-1)
        condition = (raw_cond - self.cond_mean) / self.cond_std
        for block in self.blocks:
            tokens = block(tokens, condition)
        residual = self.residual_head(self.head_norm(tokens)).squeeze(-1)
        base = self.baseline(pin, on, target_gain, target_tilt)
        prediction = base + residual
        return prediction, {"baseline": base, "residual": residual}


class SHBAwareCurvatureLoss(nn.Module):
    """MSE + penalty only for predicted curvature exceeding observed curvature + tolerance.

    Mask requires three consecutive *active* channels; true SHB holes are not
    automatically smoothed away if they are present in the training labels.
    """
    def __init__(self, lambda_curvature: float = 0.02, tolerance_db: float = 0.05):
        super().__init__()
        self.lambda_curvature = lambda_curvature
        self.tolerance_db = tolerance_db

    @staticmethod
    def d2(x):
        return x[:, 2:] - 2.0 * x[:, 1:-1] + x[:, :-2]

    def forward(self, prediction, ground_truth, active_mask):
        on = active_mask.float()
        # Zero inactive positions BEFORE any subtraction/differencing.
        # This prevents 0 * NaN from contaminating the loss on unloaded channels.
        actual = torch.where(on.bool(), ground_truth, torch.zeros_like(ground_truth))
        predicted_on = torch.where(on.bool(), prediction, torch.zeros_like(prediction))
        mse = (predicted_on - actual).square().sum() / on.sum().clamp_min(1.0)
        # SHB-aware: penalize extra/unexplained curvature, not real measured holes.
        d2_pred = self.d2(predicted_on)
        d2_true = self.d2(actual).detach()
        valid_triplets = on[:, 2:] * on[:, 1:-1] * on[:, :-2]
        excessive_curv = F.relu(d2_pred.abs() - d2_true.abs() - self.tolerance_db)
        curvature_loss = (excessive_curv.square() * valid_triplets).sum() / valid_triplets.sum().clamp_min(1.0)
        loss = mse + self.lambda_curvature * curvature_loss
        return loss, {"mse": mse.detach(), "curvature": curvature_loss.detach()}


def training_step(model, batch, optimizer, criterion, device: torch.device):
    """batch keys: pin_dbm, active_mask, aging, target_gain, target_tilt, gain_db."""
    model.train()
    b = {name: tensor.to(device) for name, tensor in batch.items()}
    optimizer.zero_grad(set_to_none=True)
    prediction, components = model(b["pin_dbm"], b["active_mask"], b["aging"],
                                   b["target_gain"], b["target_tilt"])
    loss, metrics = criterion(prediction, b["gain_db"], b["active_mask"])
    if not torch.isfinite(loss):
        raise RuntimeError("Nonfinite training loss: verify active-channel labels and scaling")
    loss.backward()
    nn.utils.clip_grad_norm_(model.parameters(), max_norm=1.0)
    optimizer.step()
    return {"loss": float(loss.detach()), **{k: float(v) for k, v in metrics.items()}}


if __name__ == "__main__":
    torch.manual_seed(2026)
    B, N = 8, 95
    # Example only: fit these two tensors on the actual training split in production.
    cond_mean = torch.tensor([0.0, 18.0, 0.0])
    cond_std = torch.tensor([1.0, 2.0, 1.0])
    net = ResidualFiLMEDFA(cond_mean, cond_std)
    loss_fn = SHBAwareCurvatureLoss(lambda_curvature=0.02, tolerance_db=0.05)
    opt = torch.optim.AdamW(net.parameters(), lr=3e-4, weight_decay=1e-4)
    batch = {
        "pin_dbm": -20.0 + torch.randn(B, N),
        "active_mask": (torch.rand(B, N) > 0.1).float(),
        "aging": torch.zeros(B),
        "target_gain": torch.full((B,), 18.0),
        "target_tilt": torch.zeros(B),
        "gain_db": 18.0 + 0.2 * torch.randn(B, N),
    }
    metrics = training_step(net, batch, opt, loss_fn, torch.device("cpu"))
    pred, parts = net(batch["pin_dbm"], batch["active_mask"], batch["aging"],
                      batch["target_gain"], batch["target_tilt"])
    assert pred.shape == (B, N) and parts["residual"].shape == (B, N)
    assert all(torch.isfinite(torch.tensor(x)) for x in metrics.values())
    print("Smoke test OK | shape:", tuple(pred.shape), "|", metrics)