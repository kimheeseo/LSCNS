"""Official OFC 2026 EDFA CSV schema to tensors; no leakage from test labels.

Age/Aging is only provided if a measured numerical operating-age feature exists.
A `Category == aging` flag is NEVER treated as an inference-available feature.
"""
from __future__ import annotations

import re
import warnings
from pathlib import Path
import numpy as np
import pandas as pd
from sklearn.model_selection import GroupShuffleSplit, train_test_split

N_CHANNELS = 95
INPUT_PREFIX = 'EDFA_input_spectra_'
MASK_PREFIX = 'DUT_WSS_activated_channel_index_'
TARGET_PREFIX = 'calculated_gain_spectra_'
AGE_CANDIDATES = ('aging_hours', 'aging_days', 'age_hours', 'age_days',
                  'EDFA_age_hours', 'device_age_hours', 'DUT_age', 'age')


def load_csv(path: Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)
    with path.open('rb') as f:
        if f.read(80).startswith(b'version https://git-lfs.github.com/spec/v1'):
            raise RuntimeError(f'Git LFS pointer: {path}. Run git lfs pull --include="Features/*/*.csv"')
    return pd.read_csv(path).reset_index(drop=True)


def ordered_cols(df: pd.DataFrame, prefix: str) -> list[str]:
    exp = re.compile('^' + re.escape(prefix) + r'(\d+)$', re.IGNORECASE)
    found = sorted((int(m.group(1)), col) for col in df.columns
                   if (m := exp.match(col)))
    if [idx for idx, _ in found] != list(range(N_CHANNELS)):
        raise ValueError(f'{prefix}: expected numbered channels 00..94; got {len(found)} columns')
    return [col for _, col in found]


def choose_age_column(df: pd.DataFrame, age_col: str | None = None) -> str | None:
    if age_col is not None:
        if age_col.lower() in ('category', 'id') or age_col.startswith('calculated_gain'):
            raise ValueError('Age conditioning must not use Category, ID, or any output label')
        if age_col not in df.columns:
            raise KeyError(f'Explicit age column not found in features: {age_col}')
        return age_col
    by_lower = {c.lower(): c for c in df.columns}
    return next((by_lower[c.lower()] for c in AGE_CANDIDATES if c.lower() in by_lower), None)


def to_arrays(features: pd.DataFrame, labels: pd.DataFrame | None = None,
              age_col: str | None = None) -> dict[str, np.ndarray]:
    pin_cols = ordered_cols(features, INPUT_PREFIX)
    mask_cols = ordered_cols(features, MASK_PREFIX)
    for required in ('target_gain', 'target_gain_tilt'):
        if required not in features.columns:
            raise ValueError(f'Missing input feature {required}')
    mask_raw = features[mask_cols].apply(pd.to_numeric, errors='coerce').fillna(0).to_numpy(np.float32)
    active_mask = (mask_raw > .5).astype(np.float32)
    pin_dbm = features[pin_cols].apply(pd.to_numeric, errors='coerce').to_numpy(np.float32)
    pin_dbm = np.nan_to_num(pin_dbm, nan=-80., posinf=20., neginf=-80.).clip(-80, 20)
    gain = pd.to_numeric(features['target_gain'], errors='coerce').to_numpy(np.float32)
    tilt = pd.to_numeric(features['target_gain_tilt'], errors='coerce').to_numpy(np.float32)
    if not np.isfinite(gain).all() or not np.isfinite(tilt).all():
        raise ValueError('Non-finite target_gain or target_gain_tilt in input features')
    age_source = choose_age_column(features, age_col)
    if age_source is None:
        age = np.zeros(len(features), dtype=np.float32)
    else:
        age = pd.to_numeric(features[age_source], errors='coerce').fillna(0).to_numpy(np.float32)
        if not np.isfinite(age).all():
            raise ValueError('Non-finite age feature')
    result = {'pin_dbm': pin_dbm, 'active_mask': active_mask, 'aging': age,
              'target_gain': gain, 'target_tilt': tilt}
    if labels is not None:
        if len(features) != len(labels):
            raise ValueError('Feature and label row counts do not match')
        if 'ID' in features and 'ID' in labels:
            if not features['ID'].astype(str).equals(labels['ID'].astype(str)):
                raise ValueError('Input and output IDs differ or are out of order')
        y = labels[ordered_cols(labels, TARGET_PREFIX)].apply(
            pd.to_numeric, errors='coerce').to_numpy(np.float32)
        if not np.isfinite(y[active_mask.astype(bool)]).all():
            raise ValueError('Active-channel labels contain NaN/Inf')
        # Labels at inactive channels are not part of loss or evaluation.
        result['gain_db'] = np.where(active_mask > 0.5, y, 0.0).astype(np.float32)
    return result


def load_official_pair(root: Path, name: str, age_col: str | None = None):
    parent = Path(root) / 'Features' / 'Train'
    if name not in ('train', 'COSMOS'):
        raise ValueError('name must be train or COSMOS')
    x = load_csv(parent / f'{name}_features.csv')
    y = load_csv(parent / f'{name}_labels.csv')
    return x, to_arrays(x, y, age_col)


def split_groups(df: pd.DataFrame, arrays: dict, val_frac: float, seed: int):
    """Keep repeated device / gain / tilt / WSS setup in the same split."""
    ids = np.arange(len(df))
    device_cols = [c for c in df if c.lower() in ('edfa_name','edfa_index', 'edfa_type','edfa_id','device')]
    devices = df[device_cols[0]].fillna('device').astype(str).to_numpy() if device_cols else ['device']*len(df)
    masks = arrays['active_mask']
    groups = [f'{devices[k]}:{arrays["target_gain"][k]:.2f}:{arrays["target_tilt"][k]:.2f}:'
              f'{np.packbits(masks[k].astype(np.uint8)).tobytes().hex()}'
              for k in range(len(df))]
    if len(set(groups)) > 1:
        splitter = GroupShuffleSplit(n_splits=1, test_size=val_frac, random_state=seed)
        it, iv = next(splitter.split(ids, groups=groups))
    else:
        warnings.warn('Only one group available; falling back to row-wise split')
        it, iv = train_test_split(ids, test_size=val_frac, random_state=seed)
    if not len(it) or not len(iv):
        raise ValueError('Empty train or validation set')
    return np.asarray(it), np.asarray(iv)


def fit_condition_stats(arrays: dict, train_indices: np.ndarray):
    cond = np.stack([arrays['aging'], arrays['target_gain'], arrays['target_tilt']], axis=1)
    mean = cond[train_indices].mean(axis=0).astype(np.float32)
    std = cond[train_indices].std(axis=0).astype(np.float32)
    # Constant condition must remain stable in FiLM. 1.0 avoids false amplification.
    std = np.where(std < 1e-3, 1.0, std)
    return mean, std


def subset_indices(indices: np.ndarray, limit: int, seed: int):
    if limit <= 0 or limit >= len(indices):
        return indices
    return np.sort(np.random.default_rng(seed).choice(indices, size=limit, replace=False))